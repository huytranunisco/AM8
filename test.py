import pandas as pd

report = pd.read_excel(r"C:\\Users\\kenguyen\\Documents\\SOPS\\NZXT\\Activity Reports\\activityReoport(NZXT - 2022-08-01 - 2022-08-15).xlsx", sheet_name="Pick Task")


group = report.groupby(by='PALLETPICKQTY').sum().reset_index
print(group)