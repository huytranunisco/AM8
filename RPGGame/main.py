import json
import bossMan
import soldierMan

boss = bossMan
sold = soldierMan

if __name__ == '__main__':


    boss.bossHPCheck() 
    print('Total HP: ', boss.bossTotalHP())

    soldierMan.getSoldierCount()
    soldierMan.getSoldierHours()
    soldierMan.getWeaponDamage()
    soldierMan.getAllSoldierWeapons()
    soldierMan.soldierAttack()

    boss.bossHPCheck() 
    print('Total HP: ', boss.bossTotalHP())
