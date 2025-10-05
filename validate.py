with open("fn-research.json", "r", encoding="utf-8") as f:
    import json
    data = json.load(f)
    print(len(data[0]["branches"]))
