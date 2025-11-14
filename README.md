提取图片EXIF坐标信息。 
支持：文件夹和子文件夹；
jpg jpeg png heic tiff bmp gif webp jfif dng cr2 nef arw rw2 sr2 raf orf pef ；
支持全格式 + 大小写自动处理 + 中文路径；
自由选择文件夹和导出文件夹保存格式有Excel格式则 csv格式。
通过gpt+Python制作。

# 支持的图片格式
IMAGE_EXTS = {
    ".jpg", ".jpeg", ".png", ".heic", ".tiff",
    ".bmp", ".gif", ".webp", ".jfif",
    ".dng", ".cr2", ".nef", ".arw", ".rw2",
    ".sr2", ".raf", ".orf", ".pef"
}

def get_exif(path):
    """获取 EXIF 数据"""
    try:
        img = Image.open(path)
        img.load()
        return img._getexif()
    except Exception as e:
        return None

def convert_to_degrees(value):
    """将 GPS 数据转为十进制度（处理分数形式）"""
    if isinstance(value, tuple):
        d, m, s = value
        return d + (m / 60) + (s / 3600)
    elif isinstance(value, (int, float)):
        return value
    elif isinstance(value, str):
        numerator, denominator = value.split('/')
        return float(numerator) / float(denominator)
    else:
        return 0

def convert_to_dms(degrees):
    """将十进制度转换为度分秒格式"""
    d = int(degrees)
    m = int((degrees - d) * 60)
    s = (degrees - d - m / 60) * 3600
    return f"{d}°{m}'{s:.4f}\""

def convert_to_dmm(degrees):
    """将十进制度转换为度分格式"""
    d = int(degrees)
    m = (degrees - d) * 60
    return f"{d}°{m:.4f}'"

def convert_to_degree_only(degrees):
    """将十进制度转换为仅度数格式（保留小数点后八位）"""
    return f"{degrees:.8f}°"

def get_gps(exif):
    """从 EXIF 中提取 GPS 坐标"""
    if exif is None:
        return None, None

    gps_info = exif.get(34853)  # 获取 GPS 信息
    if gps_info is None:
        return None, None

    try:
        lat = convert_to_degrees(gps_info[2])
        if gps_info[1] != 'N':
            lat = -lat

        lon = convert_to_degrees(gps_info[4])
        if gps_info[3] != 'E':
            lon = -lon

        return lat, lon
    except Exception as e:
        return None, None

def extract_gps(input_folder, output_folder, output_format):
    results = []

    for root, _, files in os.walk(input_folder):
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in IMAGE_EXTS:
                path = os.path.join(root, f)
                exif = get_exif(path)
                lat, lon = get_gps(exif)

                if lat is None or lon is None:
                    status = "无 GPS"
                    dms_lat = dms_lon = dmm_lat = dmm_lon = degree_only_lat = degree_only_lon = "无 GPS"
                else:
                    status = "有 GPS"
                    dms_lat = convert_to_dms(lat)
                    dms_lon = convert_to_dms(lon)
                    dmm_lat = convert_to_dmm(lat)
                    dmm_lon = convert_to_dmm(lon)
                    degree_only_lat = convert_to_degree_only(lat)
                    degree_only_lon = convert_to_degree_only(lon)

                results.append({
                    "文件路径": path,
                    "纬度 (十进制)": lat if lat is not None else "无 GPS",
                    "经度 (十进制)": lon if lon is not None else "无 GPS",
                    "纬度 (度分秒)": dms_lat,
                    "经度 (度分秒)": dms_lon,
                    "纬度 (度分)": dmm_lat,
                    "经度 (度分)": dmm_lon,
                    "纬度 (仅度数)": degree_only_lat,
                    "经度 (仅度数)": degree_only_lon,
                    "状态": status
                })

    # 导出为选定的格式
    df = pd.DataFrame(results)
    if output_format == 'Excel':
        df.to_excel(os.path.join(output_folder, "图片GPS提取结果.xlsx"), index=False)
    else:
        df.to_csv(os.path.join(output_folder, "图片GPS提取结果.csv"), index=False)
    
    messagebox.showinfo("完成", f"提取完成，结果已保存至：{output_folder}")

def select_input_folder():
    """选择输入文件夹"""
    folder = filedialog.askdirectory(title="选择输入文件夹")
    input_folder_entry.delete(0, tk.END)
    input_folder_entry.insert(0, folder)

def select_output_folder():
    """选择输出文件夹"""
    folder = filedialog.askdirectory(title="选择输出文件夹")
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, folder)

def run_extraction():
    """开始提取 GPS 数据并导出表格"""
    input_folder = input_folder_entry.get()
    output_folder = output_folder_entry.get()
    output_format = format_var.get()

    if not input_folder or not output_folder:
        messagebox.showerror("错误", "请输入有效的输入和输出文件夹路径！")
        return
    
    extract_gps(input_folder, output_folder, output_format)

# 创建 Tkinter GUI
root = tk.Tk()
root.title("图片 GPS 提取工具")
root.geometry("500x300")

# 输入文件夹选择
input_folder_label = tk.Label(root, text="输入文件夹:")
input_folder_label.pack(pady=5)
input_folder_entry = tk.Entry(root, width=50)
input_folder_entry.pack(pady=5)
input_folder_button = tk.Button(root, text="选择输入文件夹", command=select_input_folder)
input_folder_button.pack(pady=5)

# 输出文件夹选择
output_folder_label = tk.Label(root, text="输出文件夹:")
output_folder_label.pack(pady=5)
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.pack(pady=5)
output_folder_button = tk.Button(root, text="选择输出文件夹", command=select_output_folder)
output_folder_button.pack(pady=5)

# 表格格式选择
format_label = tk.Label(root, text="选择导出格式:")
format_label.pack(pady=5)
format_var = tk.StringVar(value="Excel")
format_excel = tk.Radiobutton(root, text="Excel", variable=format_var, value="Excel")
format_excel.pack(pady=5)
format_csv = tk.Radiobutton(root, text="CSV", variable=format_var, value="CSV")
format_csv.pack(pady=5)

# 确定按钮
ok_button = tk.Button(root, text="确定", command=run_extraction)  # 触发提取功能
ok_button.pack(pady=5)

# 运行 GUI
root.mainloop()
