import json
import urllib.request
import pprint
import pandas as pd


with open('C:\\Users\\kamedai\\Desktop\\python\\api_key.json') as f:
    api_key_sample = json.load(f)
    url = 'https://opendata.resas-portal.go.jp/api/v1/prefectures'
    req = urllib.request.Request(url, headers=api_key_sample)
with urllib.request.urlopen(req) as response:
    data = response.read()
print(type(data))
d = json.loads(data.decode())
#pprint.pprint(d)
df = pd.io.json.json_normalize(d['result'])
d_code = df.set_index('prefCode')['prefName'].to_dict()

s_code = df.set_index('prefCode')['prefName']
print(s_code[1])
