# 邮件客户信息提取器

这是一个自动化工具，用于从邮件对话中提取客户姓名信息。

## 功能特点

- 支持处理.txt和.docx格式的邮件文件
- 使用自然语言处理技术识别中文姓名
- 支持批量处理多个邮件文件
- 将结果导出为Excel格式

## 安装步骤

1. 安装Python 3.8或更高版本
2. 安装依赖包：
   ```bash
   pip install -r requirements.txt
   ```
3. 下载中文语言模型：
   ```bash
   python -m spacy download zh_core_web_sm
   ```

## 使用方法

1. 创建一个名为`emails`的文件夹
2. 将需要处理的邮件文件（.txt或.docx格式）放入`emails`文件夹
3. 运行程序：
   ```bash
   python email_name_extractor.py
   ```
4. 程序会自动处理所有邮件文件，并将结果保存在`extracted_names.xlsx`中

## 输出结果

程序会生成一个Excel文件，包含以下信息：
- 文件名
- 提取到的客户姓名列表

## 注意事项

- 确保邮件文件使用UTF-8编码
- 程序支持中文姓名识别
- 建议定期备份原始邮件文件
