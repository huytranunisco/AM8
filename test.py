import datetime

today = datetime.date(2022, 11, 6)
idx = (today.weekday() + 1) % 7 # MON = 0, SUN = 6 -> SUN = 0 .. SAT = 6

sun = today - datetime.timedelta(7+idx)
sat = today - datetime.timedelta(7+idx-6)
print('Last Sunday was {:%m/%d/%Y} and last Saturday was {:%m/%d/%Y}'.format(sun, sat))
'Last Sunday was 08/04/2013 and last Saturday was 08/10/2013'