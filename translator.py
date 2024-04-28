import requests
import argparse 
import winsound
import time
import datetime

# Play Windows exit sound.
winsound.PlaySound("SystemExit", winsound.SND_ALIAS)

parser = argparse.ArgumentParser() # パサー生成

parser.add_argument("source_text") # 翻訳対象のディレクトリ
parser.add_argument("source_lang") # 翻訳対象の言語
parser.add_argument("target_lang") # 翻訳後の言語
parser.add_argument("api_key") # 自身の API キー

args = parser.parse_args() # 引数解析

# パラメータの指定
params = {
            "auth_key" : args.api_key,
            "text" : args.source_text,
            "source_lang" : args.source_lang, # 翻訳対象の言語
            "target_lang": args.target_lang  # 翻訳後の言語
        }

# リクエストを投げる
request = requests.post("https://api-free.deepl.com/v2/translate", data=params) # URIは有償版, 無償版で異なるため要注意
result = request.json()["translations"][0]["text"]

print("Result:")
print(result)

with open('./translated.txt', mode='w') as f:
    f.write(result)