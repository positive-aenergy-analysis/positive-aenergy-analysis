import requests
from time import sleep
import json

vocabularies = None
with open("../vocabulary/happy.txt", "r", encoding="utf-8") as f:
    vocabularies = f.read().splitlines()
startDate = "2019-01-01"
endDate = "2019-12-31"
wholePost = "1"
final_result = []
for vocabulary in vocabularies:
    keyword = vocabulary
    URL = f"http://learn.iis.sinica.edu.tw:9187/api/sentiment?keyword={keyword}&startDate={startDate}&endDate={endDate}&wholePost={wholePost}"
    try:
        response = requests.get(URL, timeout=10)
    except requests.exceptions.Timeout as e:
        print(e)
        print(f"keyword:{vocabulary}. time out!")
        final_result.append({"keyword": keyword, "type": "do_judge", "pos": None, "neg": None})
    else:
        result = json.loads(response.text)# , encoding="utf-8"
        neg = 0
        pos = 0
        for data in result["data"]:
            if data["type"] == "neg":
                neg += data["freq"]
            elif data["type"] == "pos":
                pos += data["freq"]
        if pos > neg:
            final_result.append({"keyword": keyword, "type": "pos", "pos": pos, "neg":neg})
            print(keyword, "pos")
        elif pos < neg:
            final_result.append({"keyword": keyword, "type": "neg", "pos": pos, "neg":neg})
            print(keyword, "neg")
        else:
            final_result.append({"keyword": keyword, "type": "do_judge", "pos": pos, "neg":neg})
            print(keyword, "do_judge")
        sleep(1.5)
        
    with open("./result/test_result.txt", "w", encoding="utf-8") as f:
        json.dump(final_result, f, ensure_ascii=False, indent=2)