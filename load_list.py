import xmind

workbook = xmind.load('登录模块.xmind')
config = {
    'step': [1, -3],
    'caozuo': -3,
    'case_name': -2,
    'result': -1
}


def load_topic_info(topics, level=1, pid='root', path=''):
    topics_dict = {}
    for topic in topics:
        topics_dict[topic['id']] = {
            'id': topic['id'],
            'title': topic['title'],
            'index': level,
            'pid': pid,
            'path': f'{path}_{topic["id"]}',
            'is_last': False,
        }
        if 'topics' in topic:
            topics_dict.update(load_topic_info(topic['topics'], level=level + 1, pid=topics_dict[topic['id']]['id'],
                                               path=topics_dict[topic['id']]['path']))
        else:
            topics_dict[topic['id']]['is_last'] = True
    return topics_dict


for sheet in workbook.getSheets():
    root_topic = sheet.getRootTopic()
    topics_info = load_topic_info(root_topic.getData()['topics'])
    # print(topics_info)

    for k, v in topics_info.items():
        if v['is_last']:
            path_list = v['path'][1:].split('_')
            print('case_name', topics_info[path_list[config['case_name']]]['title'])
            step = ''

            n = 1
            for i in path_list[config['step'][0]: config['step'][1]]:
                step += f'{n}.{topics_info[i]["title"]}\n'
                n += 1
            print('step', step)
            print('caozuo', topics_info[path_list[config['caozuo']]]['title'])
            print('result', topics_info[path_list[config['result']]]['title'])
