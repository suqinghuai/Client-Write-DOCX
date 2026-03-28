import os
import shutil
import configparser
import sys

def get_script_directory():
    """获取脚本所在目录，兼容打包后的exe文件"""
    if getattr(sys, 'frozen', False):
        # 打包后的exe文件
        return os.path.dirname(sys.executable)
    else:
        # 直接运行的py文件
        return os.path.dirname(os.path.abspath(__file__))

def load_config():
    """加载配置文件"""
    script_dir = get_script_directory()
    config_path = os.path.join(script_dir, 'config3.ini')
    
    config = configparser.ConfigParser()
    
    # 尝试读取配置文件，如果格式不正确则创建默认配置
    try:
        config.read(config_path, encoding='utf-8')
    except configparser.MissingSectionHeaderError:
        # 如果配置文件没有节头，手动创建默认配置
        config['DEFAULT'] = {}
        # 尝试读取原始文件内容并解析
        with open(config_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if '=' in line:
                    key, value = line.split('=', 1)
                    config['DEFAULT'][key.strip()] = value.strip()
    
    return config

def clear_folders():
    """清空配置文件中指定的文件夹"""
    script_dir = get_script_directory()
    config = load_config()
    
    # 获取所有配置项
    folders_to_clear = []
    for key, value in config['DEFAULT'].items():
        if value.strip().lower() == 'true':
            folders_to_clear.append(key.strip())
    
    if not folders_to_clear:
        print("配置文件中没有设置为True的文件夹")
        return
    
    cleared_count = 0
    for folder_name in folders_to_clear:
        folder_path = os.path.join(script_dir, folder_name)
        
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            try:
                # 清空文件夹内容
                for filename in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, filename)
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                
                print(f"已清空文件夹: {folder_name}")
                cleared_count += 1
                
            except Exception as e:
                print(f"清空文件夹 {folder_name} 时出错: {e}")
        else:
            print(f"文件夹不存在或不是目录: {folder_name}")
    
    print(f"\n清理完成！共清空了 {cleared_count} 个文件夹")

if __name__ == "__main__":
    clear_folders()
    input("\n按回车键退出...")