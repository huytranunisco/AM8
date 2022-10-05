import json
import bigBoss
import soldier

with open('RPGGame\gameconfigs.json', 'r') as f:
    data = json.loads(f.read())

bossStats = data['boss']
weaponStats = data['weapon']
soldierStats = data['soldier']

if __name__ == '__main__':
    

    bossMan = bigBoss.BigBoss(bossStats['boss_p_hp'], bossStats['boss_m_hp'], bossStats['boss_b_hp'])