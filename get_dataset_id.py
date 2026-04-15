import requests

TOKEN = "pat_zfHHM2sTlCivOrjoqZFIw36OWSmHsf7s1t4rb6SKa1vGrVsZbdHxchd7fTVTR9at"
API_BASE = "https://api.coze.cn/v1"

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

# 获取知识库列表
response = requests.get(f"{API_BASE}/datasets", headers=headers)

if response.status_code == 200:
    datasets = response.json().get("data", [])
    print(f"找到 {len(datasets)} 个知识库：\n")
    for ds in datasets:
        print(f"名称: {ds.get('name')}")
        print(f"ID: {ds.get('dataset_id')}")
        print(f"---")
else:
    print(f"请求失败: {response.status_code}")
    print(response.text)
