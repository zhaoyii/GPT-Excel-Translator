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

直接在命令行执行便可翻译。
```
python main.py --input_table_name=example.xlsx --output_table_name=output.xlsx  --batch_size=3 --origin_language=Japanese --target_language=Chinese
```
如果在翻译过程中遇到错误，不用担心，再次执行`python main.py`即可。**程序会接着从上次翻译失败的地方开始执行**。由于 GPT 接口返回的数据，有一定概率返回错误的格式，导致程序解析错误，或者将数据输出到表格时产生错误。

**命令行参数**
- input_table_name (str): 输入Excel文件的路径和名称，默认为"example.xlsx"。
            该文件应该包含待翻译的内容。
- output_table_name (str): 输出Excel文件的路径和名称，默认为"output.xlsx"。
    该文件将包含翻译后的内容。如果文件不存在，它会被创建。
- batch_size (int): 每一批次需要翻译的记录数量，默认为3。这有助于在处理大量记录时提高效率。

- origin_language (str): 待翻译内容的原始语言，默认为"Japanese"。
- target_language (str): 目标语言，即翻译后的语言，默认为"Chinese"。

**错误日志**
- `error.log`：每次翻译后 gpt 输出的文本，记录在`error.log`中。如果发生 gpt 返回错误的文本格式，可通过`error.log`排查。

为什么要这么做？指令使 gpt 返回 json 格式的数据，这样利于程序将数据输出到表格中，保证格式正确。但 gpt 有时会有小概率返回错误的格式，所以需要记录到`error.log`中。

