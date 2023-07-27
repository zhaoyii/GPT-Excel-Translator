# GPT-Excel-Translator
Using the GPT API to translate text within Excel spreadsheets. Input a spreadsheet, output a spreadsheet.

用 GPT API 来翻译 excel 表格中的文字，输入表格，输出表格。


## Features
- 表格翻译
- 发生错误之后继续翻译

## Installation & Setup
1. 执行`poetry install`命令安装.
2. 使用`pip install pandas openai python-dotenv openpyxl`命令安装。

## Usage
**设置环境变量**
翻译需要用到 GPT API, 请在`main.py`同级目录创建一个`.env`文件，配置你的 GPT API。`.env`文件内容如下：
```
# OPENAI_MODEL 默认 gpt-3.5-turbo
OPENAI_MODEL=gpt-3.5-turbo
OPENAI_API_KEY=your_openai_api_key

# 代理配置，如果不需要可以注释掉，如果你使用了梯子，记得配置
# HTTP_PROXY=127.0.0.1:8080
# HTTPS_PROXY=127.0.0.1:8080
```

直接在命令行执行便可翻译，目前仅支持日文翻译成中文。
```
python main.py input=example.xlsx batch_size=3
```
如果在翻译过程中遇到错误，不用担心，再次执行`python main.py`即可。**程序会接着从上次翻译失败的地方开始执行**。由于 GPT 接口返回的数据，有一定概率返回错误的格式，导致程序解析错误，或者将数据输出到表格时产生错误。

**输入参数**
- `input`：输入的表格文件，默认值`example.xlsx`
- `batch_size`：每次翻译的行数，默认值`3`。建议不要设置太大，否则翻译失败浪费`token`
- `output`：输出的表格文件，默认值是`input`文件名加上`output.`前缀

