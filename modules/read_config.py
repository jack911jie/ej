import json



def read_json(fn):
    with open(fn,'r',encoding='utf8') as f:
        lines=f.readlines()

    resline=''
    for line in lines:
        newline=line.strip()
        resline=resline+newline

    config=json.loads(resline)

    return config