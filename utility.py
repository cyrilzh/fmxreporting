import os

def format_file_list_output(file_list: list) -> str:
    """
    将文件路径列表格式化为易读的多行输出。
    """
    if not file_list:
        return "（无文件被处理）"

    # 使用字典来按父目录分组
    grouped_files = {}
    
    # 查找所有路径的共同根目录，以便输出时进行简化
    common_prefix = os.path.commonprefix(file_list)
    
    for path in file_list:
        # 简化路径，移除共同前缀
        relative_path = path[len(common_prefix):].lstrip(os.path.sep)
        
        # 将文件路径拆分成目录和文件名
        directory = os.path.dirname(relative_path)
        filename = os.path.basename(relative_path)
        
        if directory not in grouped_files:
            grouped_files[directory] = []
        
        # 对临时文件（_temp_img）进行特殊标记
        if '_temp_img' in directory:
            grouped_files[directory].append(f"  └─ 🖼️ 临时文件: {filename}")
        else:
            # 原始文件
            grouped_files[directory].append(f" * 文件: {filename}")

    output_lines = []
    # 如果公共前缀有意义（不为空），先打印出来
    if common_prefix:
        output_lines.append(f"📁 根目录: {common_prefix}")
        output_lines.append("-" * 30)

    # 按目录输出
    for directory, files in grouped_files.items():
        if directory:
            output_lines.append(f"└─ 📂 文件夹: {directory}/")
            for file in files:
                output_lines.append(f"   {file}")
        else:
            # 根目录下的文件
            output_lines.append("└─ 📁 文件夹: (根目录)")
            for file in files:
                 output_lines.append(f"   {file}")
        output_lines.append("") # 目录间增加空行
            
    return "\n".join(output_lines)


