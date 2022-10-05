class Soldier:
    sTypes = ['Pierce', 'Magic', 'Blunt']
    count = 0

    def __init__(self, attackCount, dmgMod, type, soldierCost, weaponDmg, weaponCost) -> None:
        self.atkCount = attackCount
        self.dmgMod = dmgMod
        self.cost = soldierCost

        while type not in self.sTypes:
            print("Not a valid type.")
            type = input('Set type (Pierce, Magic, Blunt): ')

        self.type = type
        self.weapon = Weapon(type, weaponDmg, weaponCost)
        
        Soldier.count += 1
    
    def __str__(self) -> str:
        soldierString = f"Soldier #{self.count}"
        attributeString = "\nType: " + self.type + "\nSoldier Cost: " + str(self.cost) + "\nAmount of Turns: " + str(self.atkCount) + "\nDamage Modifier: x" + str(self.dmgMod)
        if self.weapon == None:
            weaponString = "\nNo Weapon"
        else:
            weaponString = self.weapon.__str__()
        
        return soldierString + attributeString + weaponString
    
    def attack(self):
        return self.weapon.dmg * self.dmgMod * self.atkCount

class Weapon(object):

    def __init__(self, type, wDmg, wCost) -> None:
        self.type = type
        self.dmg = wDmg
        self.cost = wCost

    def __str__(self) -> str:
        return "\nWeapon Damage: " + str(self.dmg) + "\nWeapon Cost: " + str(self.cost) + "\n"
