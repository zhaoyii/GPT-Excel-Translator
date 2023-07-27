import pandas as pd
import json
import openai
import os
from dotenv import load_dotenv, find_dotenv
from typing import List
import time
import sys
_ = load_dotenv(find_dotenv())


http_proxy = os.getenv("HTTP_PROXY")
https_proxy = os.getenv("https_proxy")
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
        ],
        request_timeout=120,
    )
    return response['choices'][0]['message']['content']


def _processing(unprocessed_records: List[dict], output_table_name, batch_size=3):
    count = 0
    
    if not output_table_name:
        raise Exception("your shoud set output_table_name")
    for i in range(0, len(unprocessed_records), batch_size):
        print(f'processed from {count} to {count+batch_size} total {len(unprocessed_records)}')

        batch = unprocessed_records[i:i+batch_size]
        chunk_str = json.dumps(list(batch), ensure_ascii=False)

        prompt = f"""
    Your task is to translate the Japanese \
    text contained within the following JSON data into Chinese，the \
    English text remains unchanged and does not require translation. \
    Please focus solely on translating the values in the JSON and ensure \
    to return your output in the correct JSON format. Specifically, your \
    returned output should be a valid JSON string, with all special characters \
    properly escaped.

    The json text delimited by triple backticks.
    ```
    {chunk_str}
    ```
    """
        content = get_completion(prompt, model)
        if content.startswith("```json"):
            content = content.lstrip("```json")
        elif content.startswith("```JSON"):
            content = content.lstrip("```JSON")
        else:
            content = content.lstrip("```")
        content = content.rstrip("```")
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
            with open('err.json', 'w', encoding='utf-8') as file:
                file.write(content)
   


def processing(input_table_name, output_table_name, batch_size=3):
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
    _processing(unprocessed_records, output_table_name, batch_size=batch_size)


if __name__ == "__main__":
    parsed_args = {}
    for arg in sys.argv[1:]:
        key, value = arg.split('=')
        parsed_args[key] = value

    start_time = time.time()
    input_table_name = parsed_args.get('input', 'example.xlsx')
    batch_size = parsed_args.get('batch_size', 3)
    output_table_name = parsed_args.get('output', "output."+input_table_name)
    print(
        f'start process at input={input_table_name} output={output_table_name} batch_size={batch_size}')
    processing(input_table_name=input_table_name,
               output_table_name=output_table_name, batch_size=int(batch_size))
    end_time = time.time()
    print(f"Function executed in {end_time - start_time:.0f} seconds")
