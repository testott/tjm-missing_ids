import re
import pandas as pd
from pandas.core.common import flatten

ycd_files = ("/top.ycd","/bot.ycd")
bom_file = "/bom.xlsx"
ref_dict = {}

for file in ycd_files:
  ycd = [line.strip() for line in open(file,'r', encoding='unicode_escape')]
  ref_ids = [line.split()[0] for line in ycd[18:-2]]

  for id in ref_ids:
    split_id = re.split('(\d+)',id)
    if split_id[0] in ref_dict:
      ref_dict[split_id[0]].append(int(split_id[1]))
    else:
      ref_dict[split_id[0]] = [int(split_id[1])]

for each in ref_dict:
  ref_dict[each].sort()

missing = {}
for each in ref_dict:
  missing_numbers = [f'{each}{str(x)}' for x in range(ref_dict[each][0], ref_dict[each][-1]+1) if x not in ref_dict[each]]
  if missing_numbers:
    missing[each] = missing_numbers

bom = pd.read_excel(bom_file,usecols = "C",skiprows = 4)
bom_ids = list(flatten(bom['REFERENCE']))

missing_from_bom = {}
missing_in_bom = {}
for each in missing:
  in_ids = []
  from_ids = []
  for id in missing[each]:
    if id in bom_ids:
      in_ids.append(id)
    else:
      from_ids.append(id)
  if in_ids:
    missing_in_bom[each] = in_ids
  if from_ids:
    missing_from_bom[each] = from_ids

print('Missing Ref IDs\n')
print('Missing From BOM:')
for each in missing_from_bom:
  print(missing_from_bom[each])
print('\nMissing, In BOM:')
for each in missing_in_bom:
  print(missing_in_bom[each])
