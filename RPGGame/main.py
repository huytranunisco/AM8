import json

if __name__ == '__main__':
    path = 'RPGGame\gameconfigs.json'

    with open(path, 'r') as f:
        data = json.loads(f.read())

    bossStats = data['boss']
    print(bossStats['boss_m_hp'])