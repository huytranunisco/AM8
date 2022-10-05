import json
import soldier
import bigBoss

if __name__ == '__main__':

    with open('RPGGame\gameconfigs.json', 'r') as f:
        data = json.loads(f.read())

    bossStats = data['boss']
    bossPHp = bossStats['boss_p_hp']
    bossMHp = bossStats['boss_m_hp']
    bossBHp = bossStats['boss_b_hp']

    weaponStats = data['weapon']
    pDmg, pCost = weaponStats['p_dmg'], weaponStats['p_cost']
    mDmg, mCost = weaponStats['m_dmg'], weaponStats['m_cost']
    bDmg, bCost = weaponStats['b_dmg'], weaponStats['b_cost']

    soldierStats = data['soldier']
    soldierCost = soldierStats['s_cost']

    bossMan = bigBoss.BigBoss(bossPHp, bossMHp, bossBHp)

    soldiersList = [[],[],[]]


    while bossMan.totalHp > 0:
        if bossMan.pHp > 0:
            s = soldier.Soldier(1, 1, 'Pierce', soldierCost, pDmg, pCost)
            bossMan.takeDamage(s.type, s.attack())
            soldiersList[0].append(s)
        if bossMan.mHp > 0:
            s = soldier.Soldier(1, 1, 'Magic', soldierCost, mDmg, mCost)
            bossMan.takeDamage(s.type, s.attack())
            soldiersList[1].append(s)
        if bossMan.bHp > 0:
            s = soldier.Soldier(1, 1, 'Blunt', soldierCost, bDmg, bCost)
            bossMan.takeDamage(s.type, s.attack())
            soldiersList[2].append(s)

    print(soldiersList)