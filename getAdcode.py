# Author: ivy
# Version: Python 3.8
# Date: 2021/03/12

# 功能：把行政区选择框内的从属关系及其对应的id爬取下来
# 输出为 adcode.json
# 需要先去”https://www.landchina.com/ExtendModule/WorkAction/EnumSelectEx.aspx?group=1&n=TAB_queryTblEnumItem_256"的源码里把第一集的名字和id存下来
# 网站中手动存下来的文件命名为 originNodes.json

import requests
import json
import random, time
import traceback
import sys


oNodes_path = "originNodes.json"

# 读取初级节点文件
def read_json(json_path):
    with open(json_path, 'r', encoding='utf8') as f:
        data = json.load(f)
        return data["zNodes"]


def req(id):
    url = "https://www.landchina.com/ExtendModule/WorkAction/EnumHandler.ashx"

    headers = {
        "accept": "text/plain, */*; q=0.01",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "content-length": "13",
        "content-type": "application/x-www-form-urlencoded",
        "cookie": "ASP.NET_SessionId=40g1s1kowx5md4035edrgv0g; Hm_lvt_83853859c7247c5b03b527894622d3fa=1615359545,1615425341,1615526595; Hm_lpvt_83853859c7247c5b03b527894622d3fa=1615526878",
        "dnt": "1",
        "origin": "https://www.landchina.com",
        "referer": "https://www.landchina.com/ExtendModule/WorkAction/EnumSelectEx.aspx?group=1&n=TAB_queryTblEnumItem_256",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36"
    }

    data = {
        "id": id,
        "group": 1
    }

    try:
        time.sleep(random.randint(1,5))
        res = requests.post(url, headers=headers, data=data)
        data = res.json()
        return data
    except Exception as e:
        print(e)
        traceback.print_exc()
        sys.exit(1)


fdata = []
oNodes = read_json(oNodes_path)
for level1_node in oNodes:
    # level 1
    level1_oname = level1_node["name"]
    level1_oid = level1_node["value"]
    print("Working on", level1_oname, "...")

    level1_tmp = {
        "name": level1_oname,
        "id": level1_oid
    }

    if level1_node["isParent"]:
        level2_nodes = req(level1_oid)

        level2_tmp_ls = []

        for level2_node in level2_nodes:
            # level 2
            level2_oname = level2_node["name"]
            level2_oid = level2_node["value"]

            print("    Working on", level2_oname, "...")

            level2_tmp = {
                "name": level2_oname,
                "id": level2_oid
            }

            if level2_node["isParent"]:
                level3_nodes = req(level2_oid)

                level3_tmp_ls = []

                for level3_node in level3_nodes:
                    print("        Working on", level3_node["name"], "...")
                    # level 3
                    level3_tmp_ls.append({
                        "name": level3_node["name"],
                        "id": level3_node["value"]
                    })
                
                level2_tmp["district"] = level3_tmp_ls
            
            level2_tmp_ls.append(level2_tmp)

        level1_tmp["city"] = level2_tmp_ls

    fdata.append(level1_tmp)

with open('adcode.json','w',encoding='utf8')as f:
    json.dump({"code": fdata}, f, ensure_ascii=False)

print("All done")