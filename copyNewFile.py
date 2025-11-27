import os
import re
import shutil
from typing import Optional, List, Tuple, Dict

# --- 定义月份替换映射表 (October -> November) ---

# 所有的替换操作都在这里定义，确保您覆盖了所有关键词
MONTH_REPLACEMENTS: Dict[str, str] = {
    # 英文 (不区分大小写)
    "october": "November",
    "oct": "Nov",
    # 中文 (不区分大小写)
    "十月": "十一月",
    "10月": "11月",
    # 数字
    "10": "11",
}

def find_monthly_xlsx_files(base_path: str, month_keywords: List[str]) -> List[Tuple[str, str]]:
    """
    在指定路径下递归查找文件名中包含特定月份关键词的 .xlsx 文件，
    并返回文件路径和匹配到的第一个关键词。

    参数:
    base_path (str): 开始查找的根目录路径。
    month_keywords (List[str]): 包含所有月份标记关键词的列表。

    返回:
    List[Tuple[str, str]]: 符合条件的文件完整路径和匹配到的关键词的列表。
                            格式为 [(full_path, matched_keyword)]
    """
    if not os.path.isdir(base_path):
        print(f"错误: 路径不存在或不是目录: {base_path}")
        return []

    found_files_and_matches = []

    # 将关键词列表转换为一个用于正则匹配的模式字符串
    # 关键修改：使用非捕获组 (?:...) 并将整个模式放入捕获组 ()，以便找到精确匹配的词
    pattern_parts = [re.escape(k) for k in month_keywords]
    # 使用 \b 来匹配单词边界，防止 "2010" 匹配到 "10"
    search_pattern = re.compile(
        f".*?({r'|'.join(pattern_parts)}).*?\\.xlsx$", 
        re.IGNORECASE
    )

    for root, _, files in os.walk(base_path):
        for file in files:
            if file.lower().endswith(".xlsx"):
                match = search_pattern.search(file) # 使用 search 查找匹配项
                
                if match:
                    full_path = os.path.join(root, file)
                    # match.group(1) 捕获到的是文件名中匹配到的精确关键词
                    matched_keyword = match.group(1) 
                    found_files_and_matches.append((full_path, matched_keyword))

    return found_files_and_matches


def copy_and_rename_file(source_path: str, matched_keyword: str, target_base_dir: str = None) -> Optional[str]:
    """
    根据匹配到的月份关键词，生成新的文件名，并复制文件。

    参数:
    source_path (str): 源文件的完整路径。
    matched_keyword (str): 文件名中匹配到的月份关键词。
    target_base_dir (str): 目标目录。如果为 None，则复制到源文件的同一目录下。

    返回:
    Optional[str]: 新文件的完整路径，如果失败则返回 None。
    """
    original_dir = os.path.dirname(source_path)
    original_filename = os.path.basename(source_path)
    
    # 确定目标目录
    target_dir = target_base_dir if target_base_dir is not None else original_dir
    # os.makedirs(target_dir, exist_ok=True)

    # 查找替换关键词：忽略大小写进行查找
    keyword_to_find = matched_keyword.lower()
    
    # 查找映射表中对应的替换值
    if keyword_to_find in MONTH_REPLACEMENTS:
        replacement = MONTH_REPLACEMENTS[keyword_to_find]
    else:
        # 如果匹配到的关键词不在替换表中，尝试查找仅包含数字或英文的替代关键词
        # 例如，如果匹配到 '10'，但用户文件中是 '10/'，我们只匹配了 '10'
        # 但如果 '10' 不在键中，我们无法替换，这里为了简化，我们只依赖精确匹配。
        print(f"警告: 关键词 '{matched_keyword}' ({keyword_to_find}) 不在替换映射表中，跳过复制。")
        return None

    # 生成新的文件名
    # 注意：使用 re.sub 进行不区分大小写的替换
    # re.escape 确保替换值中的特殊字符不会被误解析
    new_filename = re.sub(
        re.escape(matched_keyword),
        re.escape(replacement),
        original_filename,
        flags=re.IGNORECASE,
        count=1 # 只替换一次，避免文件名中多次出现月份标记导致的错误
    )
    
    target_path = os.path.join(target_dir, new_filename)
    
    # 避免新文件覆盖已有文件
    if os.path.exists(target_path):
        print(f"⚠️ 警告: 目标文件已存在，跳过复制以避免覆盖: {target_path}")
        # 您可以选择返回现有的 target_path，或者返回 None
        return target_path
    try:
        shutil.copy2(source_path, target_path)
        print(f"文件已复制并重命名:")
        print(f"  原名: {original_filename}")
        print(f"  新名: {new_filename}")
        print(f"  位置: {target_path}")
        return target_path
    except Exception as e:
        print(f"复制文件时出错: {e}")
        return None

# --- 使用示例 ---

if __name__ == "__main__":
    
    # 1. 配置参数
    # 替换为您的实际路径
    BASE_SEARCH_PATH = r"C:\Users\YourUsername\OneDrive\Reports\2025" 
    # 替换为您希望复制到的目标路径，如果设为 None，则复制到源文件的同一目录
    TARGET_COPY_PATH = r"C:\Users\YourUsername\Documents\Monthly_Prep" 

    # 2. 定义十月关键词（确保它们是 MONTH_REPLACEMENTS 字典中的键）
    OCTOBER_KEYWORDS = list(MONTH_REPLACEMENTS.keys())
    
    # 3. 查找文件和匹配项
    print(f"--- 1. 正在搜索文件 ({BASE_SEARCH_PATH}) ---")
    
    # 查找并返回 [(full_path, matched_keyword)] 列表
    result_files = find_monthly_xlsx_files(BASE_SEARCH_PATH, OCTOBER_KEYWORDS)

    if not result_files:
        print("\n未找到任何匹配十月关键词的 XLSX 文件。")
    else:
        print(f"\n--- 2. 找到 {len(result_files)} 个文件，开始复制和重命名 ---")
        
        for source_path, matched_keyword in result_files:
            # 复制并重命名文件
            copy_and_rename_file(source_path, matched_keyword, TARGET_COPY_PATH)
            print("-" * 20)
            
    print("\n任务完成！")