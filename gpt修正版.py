# -----------------------------------------------------------
# PyQt6 图片 GPS 提取工具（稳定可运行版）
# 作者：ChatGPT 修正版
# 功能：提取所有图片 GPS，支持 Excel/CSV，自定义导出名称 + 窗口图标
# -----------------------------------------------------------

import sys
import os
from pathlib import Path
from PIL import Image
import pandas as pd

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit,
    QPushButton, QFileDialog, QComboBox, QHBoxLayout, QVBoxLayout,
    QTextEdit, QProgressBar, QMessageBox, QInputDialog
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt

# 支持的图片扩展
IMAGE_EXTS = {
    ".jpg", ".jpeg", ".png", ".heic", ".tiff",
    ".bmp", ".gif", ".webp", ".jfif",
    ".dng", ".cr2", ".nef", ".arw", ".rw2",
    ".sr2", ".raf", ".orf", ".pef"
}


# ---------------------- EXIF 解析 ---------------------- #

def safe_get_exif(path: Path):
    try:
        img = Image.open(path)
        img.load()
        return img._getexif()
    except:
        return None


def convert_value(v):
    """GPS EXIF 统一转换成 float"""
    try:
        if isinstance(v, tuple):
            d, m, s = v
            return float(d) + float(m)/60 + float(s)/3600
        if isinstance(v, str) and "/" in v:
            num, den = v.split("/")
            return float(num) / float(den)
        return float(v)
    except:
        return None


def parse_lat_lon(exif):
    if not exif:
        return None, None

    gps = exif.get(34853)
    if not gps:
        return None, None

    lat = convert_value(gps.get(2))
    lon = convert_value(gps.get(4))

    if lat is None or lon is None:
        return None, None

    if gps.get(1) == "S":
        lat = -lat
    if gps.get(3) == "W":
        lon = -lon

    return lat, lon


def to_dms(val):
    d = int(abs(val))
    m = int((abs(val) - d) * 60)
    s = (abs(val) - d - m/60) * 3600
    return f"{d}°{m}'{s:.4f}\""


def to_dmm(val):
    d = int(abs(val))
    m = (abs(val) - d) * 60
    return f"{d}°{m:.6f}'"


# ---------------------- 后台线程 ---------------------- #

class Worker(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal(str)

    def __init__(self, input_folder, output_folder, fmt, out_filename):
        super().__init__()
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        self.format = fmt
        self.out_filename = out_filename

    def run(self):
        files = []

        # 递归读取图片
        for root, _, fs in os.walk(self.input_folder):
            for f in fs:
                if Path(f).suffix.lower() in IMAGE_EXTS:
                    files.append(Path(root) / f)

        if not files:
            self.finished.emit("NO_FILES")
            return

        results = []
        total = len(files)

        for i, p in enumerate(files, 1):
            exif = safe_get_exif(p)
            lat, lon = parse_lat_lon(exif)

            if lat is None:
                status = "无 GPS"
                dms_lat = dms_lon = dmm_lat = dmm_lon = deg_lat = deg_lon = "无 GPS"
            else:
                status = "有 GPS"
                dms_lat = to_dms(lat)
                dms_lon = to_dms(lon)
                dmm_lat = to_dmm(lat)
                dmm_lon = to_dmm(lon)
                deg_lat = f"{lat:.8f}°"
                deg_lon = f"{lon:.8f}°"

            results.append({
                "文件路径": str(p),
                "纬度(十进制)": lat if lat else "无 GPS",
                "经度(十进制)": lon if lon else "无 GPS",
                "纬度(度分秒)": dms_lat,
                "经度(度分秒)": dms_lon,
                "纬度(度分)": dmm_lat,
                "经度(度分)": dmm_lon,
                "纬度(仅度数)": deg_lat,
                "经度(仅度数)": deg_lon,
                "状态": status
            })

            self.progress.emit(int(i * 100 / total))
            self.log.emit(f"已处理 {i}/{total}：{p.name}")

        # 导出文件
        df = pd.DataFrame(results)
        out = self.output_folder / self.out_filename

        try:
            if self.format == "Excel":
                if not out.suffix:
                    out = out.with_suffix(".xlsx")
                df.to_excel(out, index=False)
            else:
                if not out.suffix:
                    out = out.with_suffix(".csv")
                df.to_csv(out, index=False)

            self.finished.emit(str(out))
        except Exception as e:
            self.finished.emit("ERROR:" + str(e))


# ---------------------- 主窗口 ---------------------- #

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("图片 GPS 提取工具")
        self.setWindowIcon(QIcon("app_icon.ico"))
        self.setFixedSize(700, 480)

        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        # 输入
        lbl_in = QLabel("输入文件夹：")
        self.txt_in = QLineEdit()
        btn_in = QPushButton("浏览")
        btn_in.clicked.connect(self.choose_input)

        # 输出
        lbl_out = QLabel("输出文件夹：")
        self.txt_out = QLineEdit()
        btn_out = QPushButton("浏览")
        btn_out.clicked.connect(self.choose_output)

        # 格式
        lbl_fmt = QLabel("导出格式：")
        self.cmb_fmt = QComboBox()
        self.cmb_fmt.addItems(["Excel", "CSV"])

        # 按钮
        btn_start = QPushButton("开始")
        btn_start.clicked.connect(self.start_task)

        # 进度与日志
        self.progress = QProgressBar()
        self.log = QTextEdit()
        self.log.setReadOnly(True)

        # 布局
        layout = QVBoxLayout()
        row1 = QHBoxLayout()
        row2 = QHBoxLayout()
        row3 = QHBoxLayout()

        row1.addWidget(lbl_in)
        row1.addWidget(self.txt_in)
        row1.addWidget(btn_in)

        row2.addWidget(lbl_out)
        row2.addWidget(self.txt_out)
        row2.addWidget(btn_out)

        row3.addWidget(lbl_fmt)
        row3.addWidget(self.cmb_fmt)
        row3.addStretch()
        row3.addWidget(btn_start)

        layout.addLayout(row1)
        layout.addLayout(row2)
        layout.addLayout(row3)
        layout.addWidget(self.progress)
        layout.addWidget(self.log)

        central.setLayout(layout)

    # ---------------------- UI事件 ---------------------- #

    def choose_input(self):
        d = QFileDialog.getExistingDirectory(self, "选择输入目录")
        if d:
            self.txt_in.setText(d)

    def choose_output(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.txt_out.setText(d)

    def start_task(self):
        input_folder = self.txt_in.text().strip()
        output_folder = self.txt_out.text().strip()

        if not input_folder or not output_folder:
            QMessageBox.warning(self, "错误", "请输入输入/输出目录")
            return

        # 自定义导出文件名
        default_name = Path(input_folder).name + "_GPS结果"
        name, ok = QInputDialog.getText(
            self, "导出文件名", "请输入导出文件名：", text=default_name
        )
        if not ok or not name.strip():
            return

        fmt = self.cmb_fmt.currentText()

        self.worker = Worker(input_folder, output_folder, fmt, name.strip())
        self.worker.progress.connect(self.progress.setValue)
        self.worker.log.connect(self.log.append)
        self.worker.finished.connect(self.task_finished)
        self.worker.start()
        self.log.append("开始处理...")

    def task_finished(self, msg):
        if msg == "NO_FILES":
            self.log.append("未找到任何图片。")
            return
        if msg.startswith("ERROR"):
            self.log.append("导出失败：" + msg)
            return

        self.log.append(f"完成！结果已保存至：{msg}")


# ---------------------- 入口 ---------------------- #

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
