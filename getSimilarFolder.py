import os
import difflib
from typing import List, Tuple

def find_similar_subdirs(dir1_path: str, dir2_path, folder_name_filter: str, similarity_threshold: float = 0.8) -> List[Tuple[str, str, float]]:
    """
    找出两个顶级目录中名称相似的子目录的对应关系。
    
    参数:
    dir1_path (str): 第一个顶级目录的路径。
    dir2_path (str): 第二个顶级目录的路径。
    similarity_threshold (float): 相似度阈值 (0.0 到 1.0)。只有相似度高于此阈值的子目录才会被匹配。
    
    返回:
    List[Tuple[str, str, float]]: 匹配的子目录对及其相似度百分比。
                                   格式为 [(dir1_subdir_name, dir2_subdir_name, similarity_ratio)]
    """
    
    # --- 1. 获取两个目录下的所有一级子目录名称 ---
    
    def get_first_level_subdirs(base_path: str, required_substring: str="") -> List[str]:
        """获取给定路径下的所有一级子目录名称（非嵌套）"""
        if not os.path.isdir(base_path):
            print(f"警告: 路径不存在或不是目录: {base_path}")
            return []
        
        # 使用 os.listdir 获取所有条目，并筛选出目录
        subdirs = [
            item for item in os.listdir(base_path) 
            if os.path.isdir(os.path.join(base_path, item))
        ]
        filtered_subdirs = [
            subdir for subdir in subdirs
            # 注意：这里使用了不区分大小写的查找，如果需要区分大小写，请移除 .lower()
            if required_substring.lower() in subdir.lower() 
        ]
        return filtered_subdirs

    subdirs1 = get_first_level_subdirs(dir1_path, folder_name_filter)
    subdirs2 = get_first_level_subdirs(dir2_path)
    
    if not subdirs1 or not subdirs2:
        print("至少一个目录中没有找到子目录，无法进行比较。")
        return []

    print(f"目录 '{dir1_path}' 找到子目录数量: {len(subdirs1)}")
    print(f"目录 '{dir2_path}' 找到子目录数量: {len(subdirs2)}")
    print("-" * 30)

    # --- 2. 进行相似度比较和匹配 ---
    
    matched_pairs = []
    
    # 创建一个副本，用于跟踪 dir2 中已经被匹配的子目录，避免重复匹配
    unmatched_subdirs2 = list(subdirs2) 
    
    for name1 in subdirs1:
        best_match_name = None
        max_ratio = similarity_threshold # 至少要高于阈值
        
        for name2 in unmatched_subdirs2:
            # SequenceMatcher 忽略大小写和空格会更灵活，但这里使用原始名称
            matcher = difflib.SequenceMatcher(None, name1.lower(), name2.lower())
            current_ratio = matcher.ratio()
            
            if current_ratio > max_ratio:
                max_ratio = current_ratio
                best_match_name = name2
        
        # 如果找到最佳匹配且相似度达到阈值
        if best_match_name:
            matched_pairs.append((os.path.join(dir1_path,name1), os.path.join(dir2_path,best_match_name), max_ratio))
            # 从待匹配列表中移除已匹配的子目录
            unmatched_subdirs2.remove(best_match_name)
    
    return matched_pairs

# --- 使用示例 ---

# 替换为您的实际目录路径！
DIRECTORY_A = r"C:\yy\Invoice"
DIRECTORY_B = r"C:\yy\_Landlord statements"

# 设定相似度阈值 (例如 0.8 表示至少 80% 相似)
THRESHOLD = 0.5

# 如果您想测试，可以使用下面的虚拟目录名称
# 注意：这些路径需要在您的系统上实际存在，否则 get_first_level_subdirs 会警告
# DIRECTORY_A = r"/path/to/dir_a" 
# DIRECTORY_B = r"/path/to/dir_b"

# 假设 subdirsA = ["客户A报告", "客户B报告_v1", "项目C测试"]
# 假设 subdirsB = ["客户A报告", "客户B报告_v2", "项目C文档"]

if __name__ == "__main__":
    
    # ⚠️ 实际使用时，请确保 DIRECTORY_A 和 DIRECTORY_B 是您系统上的有效路径
    
    # 这是一个虚拟的测试函数，用于演示结果
    # def mock_test():
    #     print("--- 运行虚拟测试 ---")
    #     # 模拟文件系统结构
    #     mock_dir1 = ["Project-Apple", "Project_Banana", "Test_C_Doc"]
    #     mock_dir2 = ["Project_Apple_v2", "Banana-Project", "Test-C-Docs"]

    #     subdirs1 = mock_dir1
    #     subdirs2 = list(mock_dir2) # 使用副本进行匹配

    #     matched_pairs = []
    #     THRESHOLD = 0.75 # 降低阈值以便匹配
        
    #     print(f"Dir 1: {subdirs1}")
    #     print(f"Dir 2: {subdirs2}")
    #     print(f"使用相似度阈值: {THRESHOLD}")
    #     print("-" * 30)

    #     for name1 in subdirs1:
    #         best_match_name = None
    #         max_ratio = THRESHOLD
            
    #         for name2 in subdirs2:
    #             # 注意：SequenceMatcher 比较的是 'Project-Apple' 和 'Project_Apple_v2'
    #             matcher = difflib.SequenceMatcher(None, name1.lower(), name2.lower())
    #             current_ratio = matcher.ratio()
                
    #             if current_ratio > max_ratio:
    #                 max_ratio = current_ratio
    #                 best_match_name = name2
            
    #         if best_match_name:
    #             matched_pairs.append((name1, best_match_name, max_ratio))
    #             subdirs2.remove(best_match_name) # 移除已匹配项

    #     print("\n✅ 匹配结果:")
    #     for name1, name2, ratio in matched_pairs:
    #         print(f"'{name1}' <-> '{name2}' (相似度: {ratio:.2f})")
            
    # 如果您想运行实际的文件系统扫描，请确保上面的 DIRECTORY_A 和 B 是正确的
    
    result = find_similar_subdirs(DIRECTORY_A, DIRECTORY_B, "1M" ,THRESHOLD)
    
    print("\n✅ 实际扫描匹配结果:")
    for name1, name2, ratio in result:
        if ratio > 0.7:
            print(f"目录 A: '{name1}' <-> 目录 B: '{name2}' (相似度: {ratio:.2f})")
    
