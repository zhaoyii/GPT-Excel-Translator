import pandas as pd
import json
import openai
import os
from dotenv import load_dotenv, find_dotenv
from typing import List
import time
import sys
import fire

_ = load_dotenv(find_dotenv())


http_proxy = os.getenv("HTTP_PROXY")
https_proxy = os.getenv("HTTPS_PROXY")
if http_proxy:
    os.environ["http_proxy"] = http_proxy
if https_proxy:
    os.environ["https_proxy"] = https_proxy

openai.api_key = os.getenv('OPENAI_API_KEY')
# gpt-3.5-turbo 或者 gpt-4
openai_model = os.getenv('OPENAI_MODEL')
model = 'gpt-3.5-turbo'
if openai_model:
    model = openai_model


def read_excel(table_nanme) -> List[dict]:
    data = pd.read_excel(table_nanme, engine='openpyxl')
    # 将 NaN 值替换为 None
    data = data.where(pd.notnull(data), None)
    # 将 DataFrame 转换为一个字典列表
    data_records = data.to_dict(orient="records")
    return data_records


def get_completion(prompt, model="gpt-3.5-turbo") -> str:
    response = openai.ChatCompletion.create(
        model=model,
        messages=[
            {"role": "system", "content": "You are a skilled translator."},
            {"role": "user", "content": prompt},
        ]
    )
    return response['choices'][0]['message']['content']


def _translate(unprocessed_records: List[dict], output_table_name, batch_size, origin_language, target_language):
    count = 0

    if not output_table_name:
        raise Exception("your shoud set output_table_name")
    for i in range(0, len(unprocessed_records), batch_size):
        print(
            f'processed from {count} to {count+batch_size} total {len(unprocessed_records)}')

        batch = unprocessed_records[i:i+batch_size]
        chunk_str = json.dumps(list(batch), ensure_ascii=False)

        prompt = f"""
    Your task is to translate the {origin_language} \
    text contained within the following JSON data into {target_language}，the \
    English text remains unchanged and does not require translation. \
    Please focus solely on translating the values in the JSON and ensure \
    to return your output in the correct JSON format. 
    
    Specifically, your \
    returned output should be a valid JSON string, with all special characters \
    properly escaped. Like follow:
    ```
    [{{"A": 1, "B": 1}},{{"A": 2, "B": 2}},{{"A": 3, "B": 3}}]
    ```

    The json text delimited by triple backticks.
    ```
    {chunk_str}
    ```
    """
        raw_content = get_completion(prompt, model)
        content = extract_content(raw_content)
        try:
            contents = json.loads(content)
            df_to_append = pd.DataFrame(contents)
            with pd.ExcelFile(output_table_name) as reader:
                existing_data = pd.read_excel(reader, sheet_name='Sheet1')
            appended_data = pd.concat(
                [existing_data, df_to_append], ignore_index=True)

            with pd.ExcelWriter(output_table_name, engine='openpyxl', mode='a') as writer:
                writer.book.remove(writer.book['Sheet1'])
                appended_data.to_excel(
                    writer, sheet_name='Sheet1', index=False)
            count += batch_size
        except Exception as e:
            print(f"An exception occurred: {e}")
            raise e
        finally:
            with open('error.log', 'w', encoding='utf-8') as file:
                file.write(content)


def extract_content(content: str) -> str:
    if content.startswith("```json"):
        content = content.lstrip("```json")
    elif content.startswith("```JSON"):
        content = content.lstrip("```JSON")
    else:
        content = content.lstrip("```")
    content = content.rstrip("```")
    return content


def translate(
        input_table_name: str = "example.xlsx",
        output_table_name: str = "output.xlsx",
        batch_size: int = 3,
        origin_language: str = "Japanese",
        target_language: str = "Chinese"
):
    """
    Reads content to be translated from a specified Excel file and saves the translated results into another Excel file.

    If the output file already exists, the function will start translating from the last processed record to ensure no duplication of processing.

    Parameters:
        input_table_name (str): Path and name of the input Excel file, default is "example.xlsx".
            This file should contain the content to be translated.

        output_table_name (str): Path and name of the output Excel file, default is "output.xlsx".
            This file will contain the translated content. If the file does not exist, it will be created.

        batch_size (int): Number of records to be translated in each batch, default is 3. This helps improve efficiency when dealing with a large number of records.

        origin_language (str): Original language of the content to be translated, default is "Japanese".

        target_language (str): Target language, i.e., the language after translation, default is "Chinese".

    Returns:
        None. After the function is executed, all translated content will be saved in the Excel file specified by output_table_name.

    Example:
        translate("input_data.xlsx", "translated_data.xlsx", batch_size=5, origin_language="English", target_language="Chinese")

    Notes:
        - Ensure the required libraries, such as pandas, are installed.
        - Input and output files should be in the same directory unless a full path is provided.
        - The function relies on a helper function named `_translate` to perform the actual translation operation.
    """

    print(f'start process at input={input_table_name} output={output_table_name} \
          batch_size={batch_size} origin_language={origin_language} target_language={target_language}')

    input_records = read_excel(input_table_name)
    output_records = None
    if os.path.exists(output_table_name):
        output_records = read_excel(output_table_name)
    else:
        df = pd.DataFrame()
        df.to_excel(output_table_name, index=False)
    start_row = 0
    if output_records:
        start_row = len(output_records)
    print("start at row ", start_row)
    unprocessed_records = input_records[start_row:]
    _translate(unprocessed_records, output_table_name,
               batch_size, origin_language, target_language)


if __name__ == "__main__":
    start_time = time.time()

    fire.Fire(translate)

    end_time = time.time()
    print(f"Function executed in {end_time - start_time:.0f} seconds")
