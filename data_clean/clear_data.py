import json
test_results = None
final_vocabularies = []
do_judge = []
with open("./result/test_result.txt", "r", encoding="utf-8") as f:
    test_results = json.load(f)



for test_result in test_results:
    if test_result["type"] == "pos":
        final_vocabularies.append(test_result["keyword"])
    elif test_result["type"] == "do_judge" and test_result["pos"] != None and test_result["pos"] != 0:
        do_judge.append(test_result)

vocabularies = "\n".join(final_vocabularies)
with open("./result/final_result.txt", "w", encoding="utf-8") as f:
    f.write(vocabularies)
with open("./result/do_judge.txt", "w", encoding="utf-8") as f:
    json.dump(do_judge, f, ensure_ascii=False, indent=2)