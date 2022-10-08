# dict and inputs
d = {'money':100}
soldierInput = int(input('Input soldier ammount to use this turn '))
swordInput = int(input('Input sword ammount to use this turn '))
maceInput = int(input('Input mace ammount to use this turn '))
staffInput = int(input('Input staff ammount to use this turn '))
x = 0
t = 0
u = 0
v = 0
# function for soldier, cost 16g per call/turn
def soldierCost():
    for money in d:
      d[money] = d[money] - 16
      return print('Total gold left after soldier cost: ', (d[money]))

# function for sword cost, 1g per call/turn
def swordCost():
  for money in d:
    d[money] = d[money] - 1
    return print('Total gold left after sword cost: ', (d[money]))

# function for mace cost, 12g per call/turn
def maceCost():
  for money in d:
    d[money] = d[money] - 12
    return print('Total gold left after mace cost: ', (d[money]))

# function for staff cost, 3g per call/turn
def staffCost():
  for money in d:
    d[money] = d[money] - 3
    return print('Total gold left after staff cost: ', (d[money]))

# while loops for inputs

while x < soldierInput:
  soldierCost()
  x = x + 1
while t < swordInput:
  swordCost()
  t = t + 1
while u < maceInput:
  maceCost()
  u = u + 1
while v < staffInput:
  staffCost()
  v = v + 1

# total gold remaining after turn
print('Total remaining gold after turn:', d['money'])