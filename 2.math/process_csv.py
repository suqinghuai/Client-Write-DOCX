import pandas as pd
import os
import glob
import sys
from pathlib import Path
import msvcrt  # Windows平台按键检测

def calculate_rebar_measurement(encrypted_length, encrypted_count, non_encrypted_length, non_encrypted_count):
    """
    计算箍筋实测值
    实际执行除法运算并显示结果
    """
    try:
        # 转换为数值类型
        encrypted_len = float(encrypted_length) if encrypted_length else 0
        encrypted_cnt = float(encrypted_count) if encrypted_count else 0
        non_encrypted_len = float(non_encrypted_length) if non_encrypted_length else 0
        non_encrypted_cnt = float(non_encrypted_count) if non_encrypted_count else 0
        
        # 避免除零错误
        if encrypted_cnt == 0:
            encrypted_result = 0
        else:
            encrypted_result = encrypted_len / encrypted_cnt
            
        if non_encrypted_cnt == 0:
            non_encrypted_result = 0
        else:
            non_encrypted_result = non_encrypted_len / non_encrypted_cnt
        
        # 四舍五入到整数
        encrypted_result = round(encrypted_result)
        non_encrypted_result = round(non_encrypted_result)
        
        # 返回实际运算结果的字符串
        return f"@{non_encrypted_result}（{encrypted_result}）"
    except (ValueError, TypeError):
        # 如果数据无法转换为数字，返回默认格式
        return "  "

def clean_csv_data(file_path):
    """清洗CSV文件数据，去除Markdown格式残留"""
    try:
        # 读取原始文件内容
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # 过滤掉包含Markdown分隔符的行（如---,或:--,等）
        cleaned_lines = []
        for line in lines:
            # 跳过包含Markdown表格分隔符的行
            if '---,' in line or ':--,' in line or line.strip().startswith('---'):
                continue
            cleaned_lines.append(line)
        
        # 如果没有数据行，返回空DataFrame
        if len(cleaned_lines) <= 1:  # 只有表头或没有数据
            return pd.DataFrame()
        
        # 将清洗后的内容写回临时文件或直接解析
        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as temp_file:
            temp_file.writelines(cleaned_lines)
            temp_path = temp_file.name
        
        # 读取清洗后的CSV文件
        df = pd.read_csv(temp_path)
        
        # 清理临时文件
        import os
        os.unlink(temp_path)
        
        return df
        
    except Exception as e:
        print(f"清洗文件 {file_path} 时出错：{e}")
        return pd.DataFrame()

def process_csv_file(file_path):
    """处理单个CSV文件"""
    try:
        # 先进行数据清洗
        df = clean_csv_data(file_path)
        
        # 检查是否成功读取数据
        if df.empty:
            print(f"警告：文件 {file_path} 清洗后无有效数据，跳过处理")
            return False
        
        # 检查必要的列是否存在
        required_columns = ['加密区箍筋长度(mm)', '加密区箍筋档数', '非加密区箍筋长度(mm)', '非加密区箍筋档数']
        for col in required_columns:
            if col not in df.columns:
                print(f"警告：文件 {file_path} 中缺少列 '{col}'，跳过处理")
                return False
        
        # 新增箍筋实测列
        df['箍筋实测'] = df.apply(
            lambda row: calculate_rebar_measurement(
                row['加密区箍筋长度(mm)'], 
                row['加密区箍筋档数'],
                row['非加密区箍筋长度(mm)'],
                row['非加密区箍筋档数']
            ), 
            axis=1
        )
        
        # 获取当前工作目录（打包后也能正确工作）
        current_dir = Path.cwd()
        success_folder = current_dir / 'CSV数据处理结果'
        success_folder.mkdir(exist_ok=True)
        
        # 生成新的文件名（原文件名称+p）
        file_name = Path(file_path).name
        new_file_name = file_name.replace('.csv', 'p.csv')
        output_path = success_folder / new_file_name
        
        # 保存处理后的文件到CSV数据处理结果文件夹
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        
        print(f"成功处理文件：{file_path}")
        print(f"输出文件：{output_path}")
        print(f"处理了 {len(df)} 行数据")
        
        return True
        
    except Exception as e:
        print(f"处理文件 {file_path} 时出错：{e}")
        return False

def batch_process_csv_files(folder_path):
    """批量处理文件夹中的所有CSV文件"""
    # 将相对路径转换为绝对路径
    folder_path = Path(folder_path).resolve()
    
    if not folder_path.exists():
        print(f"错误：文件夹 '{folder_path}' 不存在")
        return
    
    # 查找所有CSV文件
    csv_files = list(folder_path.glob("*.csv"))
    
    if not csv_files:
        print(f"在文件夹 {folder_path} 中未找到CSV文件")
        return
    
    print(f"找到 {len(csv_files)} 个CSV文件")
    
    success_count = 0
    
    for csv_file in csv_files:
        # 跳过已处理的文件（避免重复处理）
        if 'p.csv' in csv_file.name:
            continue
            
        if process_csv_file(str(csv_file)):
            success_count += 1
    
    print(f"\n批量处理完成！")
    print(f"成功处理：{success_count} 个文件")
    print(f"失败：{len(csv_files) - success_count} 个文件")

def get_resource_path(relative_path):
    """获取资源文件的绝对路径，支持打包环境"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = Path.cwd()
    
    return Path(base_path) / relative_path

def wait_for_key_press():
    """等待用户按任意键退出"""
    print("\n程序执行完毕，按任意键退出...")
    # 等待用户按键
    msvcrt.getch()

if __name__ == "__main__":
    # 设置CSV文件所在文件夹路径
    csv_folder = "转CSV结果"
    
    # 获取当前工作目录
    current_dir = Path.cwd()
    print(f"当前工作目录：{current_dir}")
    
    # 检查文件夹是否存在
    csv_folder_path = current_dir / csv_folder
    if not csv_folder_path.exists():
        print(f"错误：文件夹 '{csv_folder_path}' 不存在")
        print("请确保csv_output文件夹与程序在同一目录下")
    else:
        # 批量处理CSV文件
        batch_process_csv_files(str(csv_folder_path))
    
    # 等待用户按键退出
    wait_for_key_press()