import os
from typing import Dict
from pathlib import Path

from getSimilarFolder import find_similar_subdirs 
from copyNewFile import copy_and_rename_file, find_monthly_xlsx_files, MONTH_REPLACEMENTS
from insertPic2Excel import insert_images_to_excel_with_pdf
from utility import format_file_list_output

# -----------------------------------------------------------
# I. 辅助变量定义
# -----------------------------------------------------------

# 替换映射表 (从您的代码中复制)
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.pdf') # 包含 .pdf
FOLD_NAME_FILTER ="NOV"

# -----------------------------------------------------------
# II. 主控函数
# -----------------------------------------------------------

def automate_monthly_report_prep_final(base_report_dir: str, base_image_dir: str, similarity_threshold: float = 0.7) -> None:
    """
    协调整个任务流程的主函数，第一步直接调用 find_similar_subdirs 获取目录映射。

    参数:
    base_report_dir (str): 包含 Excel 报告子文件夹的根目录。
    base_image_dir (str): 包含图片子文件夹的根目录。
    similarity_threshold (float): 查找相似图片目录的相似度阈值。
    """
    
    if not os.path.isdir(base_report_dir):
        print(f"❌ 错误: 报告基础路径不存在: {base_report_dir}")
        return
    if not os.path.isdir(base_image_dir):
        print(f"❌ 错误: 图片基础路径不存在: {base_image_dir}")
        return

    # --- 1. 直接调用 find_similar_subdirs 获取报告-图片映射清单 ---
    
    # 假设 find_similar_subdirs 内部逻辑会比较 base_report_dir 和 base_image_dir 
    # 下的所有子目录名称，并返回 [(ReportPath, ImagePath, Ratio)] 列表。
    
    try:
        # 注意：此处假设您已在其他地方定义和导入 find_similar_subdirs
        # 实际运行中，您需要确保 find_similar_subdirs 是可用的。
        directory_mappings = find_similar_subdirs(base_report_dir, base_image_dir,FOLD_NAME_FILTER, similarity_threshold)
    except NameError:
        print("❌ 错误: 找不到 find_similar_subdirs 函数的定义。请确保其已被导入或定义。")
        return
    except Exception as e:
        print(f"❌ 错误: 调用 find_similar_subdirs 时发生错误: {e}")
        return

    if not directory_mappings:
        print(f"❌ 错误: 在相似度 > {similarity_threshold} 的阈值下，未找到任何报告目录与图片目录的匹配对。任务中止。")
        return

    print(f"--- 1. 成功匹配 {len(directory_mappings)} 对报告/图片文件夹 ---")
    print("-" * 50)
    
    total_processed_folders = 0

    # --- 2. 循环遍历映射并执行复制/插入操作 ---
    all_images_processed=[]
    for current_report_folder_path, best_match_image_path, ratio in directory_mappings:
        report_folder_name = Path(current_report_folder_path).name
        image_folder_name = Path(best_match_image_path).name
        
        print(f"\n>>>> 正在处理报告文件夹: {report_folder_name} (图片源: {image_folder_name}, 相似度: {ratio:.2f}) <<<<")

        # --- 2.1 查找 10 月份的 Excel 文件 ---
        report_files = find_monthly_xlsx_files(current_report_folder_path)
        
        if not report_files:
            print("❌ 警告: 未在当前报告子文件夹中找到 10 月份的 XLSX 文件，跳过。")
            continue

        # 假设我们只对每个子文件夹中找到的第一个 10 月文件进行操作
        source_excel_path, matched_keyword = report_files[0]
        print(f"✅ 找到源 Excel: {os.path.basename(source_excel_path)}")

        # --- 2.2 复制并重命名为 11 月版本 ---
        # current_report_folder_path=r"c:\yy\test"
        current_report_folder_path=os.path.dirname(source_excel_path)
        new_excel_path = copy_and_rename_file(source_excel_path, matched_keyword, current_report_folder_path)
        
        if not new_excel_path:
            print("❌ 错误: 文件复制或重命名失败，跳过后续步骤。")
            continue
            
        print(f"✅ 创建 11 月文件: {os.path.basename(new_excel_path)}")

        # --- 2.3 收集图片和 PDF 文件路径 ---
        
        image_and_pdf_files = []
        try:
            # 遍历 best_match_image_path 目录下的所有条目
            for item_name in os.listdir(best_match_image_path):
                file_path = os.path.join(best_match_image_path, item_name)
                
                # 1. 检查条目是否是文件 (排除子目录)
                if os.path.isfile(file_path):
                    ext = os.path.splitext(item_name)[1].lower()
                    
                    # 2. 检查扩展名是否在允许的列表中
                    if ext in IMAGE_EXTENSIONS:
                        image_and_pdf_files.append(file_path)
                        all_images_processed.append(file_path)
        except FileNotFoundError:
            print(f"警告: 目录未找到 - {best_match_image_path}")
        except Exception as e:
            print(f"在收集文件时发生错误: {e}")

        if not image_and_pdf_files:
            print("❌ 警告: 图片文件夹中未找到任何图片或 PDF 文件。")
        else:
            print(f"✅ 找到 {len(image_and_pdf_files)} 个图片/PDF 文件。")

            # --- 2.4 查找 Excel 目标位置和工作表名称 ---
            
            # --- 2.5 插入图片和 PDF ---
            insert_images_to_excel_with_pdf(
                excel_path=new_excel_path,
                file_paths=image_and_pdf_files,
            )
            print(f"✅ 图片/PDF 插入完成。")

        total_processed_folders += 1
        print("<<<< 当前文件夹处理完毕 >>>>")
        
    print(f"\n\n🎉🎉 自动化流程全部完成！总共处理了 {total_processed_folders} 个文件夹。 🎉🎉")
    formatted_output = format_file_list_output(all_images_processed)
    print("--- 所有被处理的图片/PDF文件列表 ---")
    print(formatted_output)


# --- 执行示例 ---
if __name__ == "__main__":
    
    # ⚠️ 替换为您的实际路径，确保这两个路径都存在
    REPORT_FOLDER = r"C:\yy\_Landlord statements" # 报告根目录 (包含 Project A, Project B等子文件夹)
    IMAGE_FOLDER = r"C:\yy\Invoice" # 图片根目录 (包含 Project A Photos, Project B Pics等子文件夹)

    # 请确保您已经定义或导入了所有的辅助函数，否则代码会因 NameError 而失败。

    automate_monthly_report_prep_final(
        base_report_dir=REPORT_FOLDER,
        base_image_dir=IMAGE_FOLDER,
    )
    
    # print("\n代码已修正为直接调用 find_similar_subdirs。请在实际运行前确保所有导入和函数定义都是完整的。")