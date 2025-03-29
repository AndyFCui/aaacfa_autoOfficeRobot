import spacy
import pandas as pd
import os
import re
from typing import List, Dict, Tuple

class EmailNameExtractor:
    def __init__(self):
        # 加载中文语言模型
        self.nlp = spacy.load("zh_core_web_sm")
        
    def extract_names_from_text(self, text: str) -> List[str]:
        """从文本中提取人名"""
        doc = self.nlp(text)
        names = []
        
        # 使用NER识别人名
        for ent in doc.ents:
            if ent.label_ == "PERSON":
                names.append(ent.text)
        
        # 使用正则表达式匹配常见的中文姓名模式
        # 更新姓氏列表，只包含最常见的姓氏
        name_pattern = r'(?:(?:张|王|李|赵|刘|陈|杨|黄|周|吴|徐|孙|马|朱|胡|郭|何|高|林|罗|郑|梁|谢|宋|唐|许|韩|冯|邓|曹|彭|曾|肖|田|董|袁|潘|于|蒋|蔡|余|杜|叶|程|苏|魏|吕|丁|任|沈|姚|卢|姜|崔|钟|谭|陆|汪|范|金|石|廖|贾|夏|郎|方|侯|邹|熊|孔|秦|白|江|阎|薛|尹|段|雷|黎|史|龙|贺|陶|顾|毛|郝|龚|邵|万|钱|严|赖|覃|洪|武|莫|孟)(?:[一-龥]){1,2})(?:(?:先生|小姐|女士|经理|总监|老师|工程师|总经理|董事长|主任|总|工))?'
        matches = re.findall(name_pattern, text)
        names.extend(matches)
        
        # 去重并过滤掉明显的误识别
        filtered_names = []
        stop_words = {'先生', '小姐', '女士', '经理', '总监', '老师', '工程师', '总经理', '董事长', '主任', '总', '工'}
        for name in set(names):
            # 移除职位头衔
            for title in stop_words:
                name = name.replace(title, '')
            # 过滤掉长度不合适的名字
            if 2 <= len(name) <= 3 and not any(word in name for word in ['我们', '你们', '他们', '什么', '为什', '那些', '这些', '任何', '所有']):
                filtered_names.append(name)
        
        return list(set(filtered_names))
    
    def process_email_file(self, file_path: str) -> Dict[str, List[str]]:
        """处理邮件文件并提取人名"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
            
            names = self.extract_names_from_text(text)
            return {
                'file_name': os.path.basename(file_path),
                'names': names
            }
        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {str(e)}")
            return {
                'file_name': os.path.basename(file_path),
                'names': []
            }
    
    def process_directory(self, directory_path: str) -> pd.DataFrame:
        """处理整个目录下的邮件文件"""
        results = []
        
        for filename in os.listdir(directory_path):
            if filename.endswith('.txt'):
                file_path = os.path.join(directory_path, filename)
                result = self.process_email_file(file_path)
                results.append(result)
        
        return pd.DataFrame(results)

def main():
    # 创建提取器实例
    extractor = EmailNameExtractor()
    
    # 设置邮件目录路径
    email_dir = "emails"  # 请确保这个目录存在并包含邮件文件
    
    # 处理所有邮件文件
    df = extractor.process_directory(email_dir)
    
    # 保存结果到Excel文件
    df.to_excel("extracted_names.xlsx", index=False)
    print("处理完成！结果已保存到 extracted_names.xlsx")
    
    # 打印结果
    print("\n提取结果：")
    for _, row in df.iterrows():
        print(f"\n文件：{row['file_name']}")
        print(f"提取到的姓名：{', '.join(row['names'])}")

if __name__ == "__main__":
    main() 