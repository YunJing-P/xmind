import xmind

workbook = xmind.load('进阶测试.xmind')


def load_case_path(data, begin=0, pid='root', path=''):
    new_dict = {}
    new_list = []
    for i in data:
        new_dict[i['id']] = {
            'id': i['id'],
            'title': i['title'],
            'index': begin,
            'pid': pid,
            'path': f'{path}_{i["id"]}'
        }
        new_list.append(new_dict[i['id']])
        if 'topics' in i:
            # print(i['topics'])
            new_dict[i['id']]['child'], nnew_list = load_case_path(i['topics'], begin + 1, i['id'], new_dict[i['id']]['path'])
            new_list += nnew_list
        else:
            new_dict[i['id']]['child'] = None
    return new_dict, new_list


for sheet in workbook.getSheets():
    root_topic = sheet.getRootTopic()
    a, b = load_case_path(root_topic.getData()['topics'])
    # print(a)
    for i in b:
        if i['child'] is None:
            print(i['title'], i['path'])
    # print(root_topic.getData()['topics'])
