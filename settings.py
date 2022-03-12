import re

WIDTH, HEIGHT = 860, 300
POS_X, POS_Y = 1100, 400

# REGEX
id_regex = re.compile(r'\d{1,4} ')
item_regex = re.compile(r' \d{6,7}')
quantity_regex = re.compile(r'(?<=2022 )\w+')

# File names
generated_txt = "data.txt"
log_txt = "log.txt"
generated_bom = "kusovnik.xlsx"
