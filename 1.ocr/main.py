import os
import base64
import configparser
from pathlib import Path
from openai import OpenAI
import time
import json
import subprocess
import logging
import csv
import re

class TableOCRProcessor:
    def __init__(self, config_path="config.ini"):
        self.config = configparser.ConfigParser()
        self.config.read(config_path, encoding='utf-8')
        
        # 初始化OpenAI客户端
        self.client = OpenAI(
            api_key=self.config['API']['api_key'],
            base_url=self.config['API']['base_url']
        )
        
        # 创建输出目录
        self.image_folder = Path(self.config['Paths']['image_folder'])
        self.output_folder = Path(self.config['Paths']['output_folder'])
        self.output_folder.mkdir(exist_ok=True)
        
        # 提示词模板
        self.prompt_template = """这是一张混凝土梁、柱截面尺寸及配筋检测原始记录的照片，请完整提取表格内容，严格按照原表列名输出Markdown表格，不要遗漏任何数据，包括备注里的测点编号和坐标。

列名：序号、构件类型、构件位置、截面尺寸(mm)、截面型式示意、纵筋规格、纵筋数量、箍筋规格、非加密区箍筋长度(mm)、非加密区箍筋档数、加密区箍筋长度(mm)、加密区箍筋档数、加密区长度(mm)、保护层厚度(mm)、备注

请确保：
1. 表格格式正确，使用Markdown表格语法
2. 注意**重点识别所有分数（如1/4、2/3），必须以a/b格式输出，不得转为0.25、0.666等小数**，保持表格行列对齐。
3. 所有数据准确无误
4. 不要遗漏任何行或列的数据
5. 如果图片中有多个表格，请全部提取"""

        # 初始化日志记录
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / f"log_{time.strftime('%Y%m%d_%H%M%S')}.txt"
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        logging.info("程序启动")

    def image_to_base64(self, image_path):
        """将图片转换为base64编码"""
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    def process_single_image(self, image_path):
        """处理单张图片"""
        logging.info(f"正在处理图片: {image_path.name}")
        try:
            # 将图片转换为base64
            base64_image = self.image_to_base64(image_path)
            
            # 构建消息
            messages = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": self.prompt_template
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ]
            
            # 调用API
            response = self.client.chat.completions.create(
                model=self.config['API']['model'],
                messages=messages,
                max_tokens=4000,
                temperature=0.1
            )
            
            result = response.choices[0].message.content
            logging.info(f"图片 {image_path.name} 处理成功")
            return result
            
        except Exception as e:
            logging.error(f"处理图片 {image_path.name} 时出错: {str(e)}")
            return None

    def save_result(self, image_name, result):
        """保存识别结果并自动转换为CSV"""
        logging.info(f"正在保存图片 {image_name} 的识别结果")
        # 生成输出文件名
        output_filename = f"{Path(image_name).stem}_识别结果.md"
        output_path = self.output_folder / output_filename

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"# 图片识别结果: {image_name}\n\n")
            f.write("## 识别时间: " + time.strftime("%Y-%m-%d %H:%M:%S") + "\n\n")
            f.write("## 提取的表格内容:\n\n")
            f.write(result)

        print(f"结果已保存到: {output_path}")

        # 自动转换为CSV格式
        try:
            self.convert_to_csv()
            logging.info("识别结果已转换为CSV格式")
        except Exception as e:
            logging.error(f"转换为CSV时出错: {e}")

        return output_path

    def process_all_images(self):
        """处理图片文件夹中的所有图片"""
        logging.info("开始处理所有图片")
        # 支持的图片格式
        image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif'}
        
        # 获取所有图片文件
        image_files = [f for f in self.image_folder.iterdir() 
                      if f.is_file() and f.suffix.lower() in image_extensions]
        
        if not image_files:
            print(f"在目录 {self.image_folder} 中没有找到图片文件")
            return
        
        print(f"找到 {len(image_files)} 个图片文件，开始处理...")
        
        results = []
        for image_file in image_files:
            # 处理图片
            result = self.process_single_image(image_file)
            
            if result:
                # 保存结果
                output_path = self.save_result(image_file.name, result)
                results.append({
                    'image': image_file.name,
                    'output': output_path,
                    'result': result
                })
                
                # 添加延迟避免API限制
                time.sleep(2)
            else:
                print(f"图片 {image_file.name} 处理失败")
        
        logging.info(f"成功处理 {len(results)} 张图片")
        return results

    def convert_to_csv(self):
        """
        将OCR结果转换为CSV格式
        """
        input_directory = "转md文档结果"
        output_directory = "转CSV结果"
        
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        for filename in os.listdir(input_directory):
            if filename.endswith("_识别结果.md"):
                input_path = os.path.join(input_directory, filename)
                output_path = os.path.join(output_directory, filename.replace("_识别结果.md", ".csv"))

                with open(input_path, "r", encoding="utf-8") as infile:
                    lines = infile.readlines()

                table_content = []
                for line in lines:
                    # 使用正则表达式识别表格内容
                    if re.match(r"^\|.*\|$", line.strip()):
                        table_content.append([cell.strip() if cell.strip() else "" for cell in line.strip().split("|")])

                # 移除空行
                cleaned_table = [
                    row for row in table_content if any(cell for cell in row)
                ]

                # 写入CSV
                with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerows(cleaned_table)

def main():
    """主函数"""
    print("=== 手写表格识别程序启动 ===")
    print("开发者联系方式：430615396@qq.com")

    # 检查配置文件
    if not os.path.exists("config.ini"):
        print("错误: 未找到config.ini配置文件")
        print("请确保配置文件存在并已正确配置API密钥")
        input("按任意键退出...")
        return

    try:
        # 创建处理器实例
        processor = TableOCRProcessor()

        # 处理所有图片
        results = processor.process_all_images()

        if results:
            print(f"\n=== 处理完成 ===")
            print(f"成功处理 {len(results)} 张图片")
            print(f"结果保存在: {processor.output_folder}")
            print("所有识别结果已成功转换为CSV格式")
        else:
            print("没有成功处理的图片")

    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        print("请检查配置文件中的API密钥和网络连接")

    input("按任意键退出...")

if __name__ == "__main__":
    main()