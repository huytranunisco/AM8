class BigBoss:
    def __init__(self, pierceHp, magicHp, bluntHp) -> None:
        self.pHp = pierceHp
        self.mHp = magicHp
        self.bHp = bluntHp
        self.totalHp = self.pHp + self.mHp + self.bHp

    def getHPs(self):
        stringOutput = 'Total HP = ' + str(self.totalHp) + '\nPierce HP = ' + str(self.pHp) + '\nMagic HP = ' + str(self.mHp) + '\nBlunt HP = ' + str(self.bHp)
        return stringOutput
    
    def takeDamage(self, type, dmg):
        if type == 'Pierce':
            self.pHp -= dmg
        if type == 'Magic' or type == 'Blunt':
            self.mHp -= dmg
        if type == 'Blunt':
            self.bHp -= dmg
        
        self.totalHp -= dmg