import openpyxl as xl
from openpyxl.styles import Font

MAINFILE = 'WR/8.0_Stats_v3.0_20220507.xlsx'
wb = xl.load_workbook(MAINFILE)
print(wb.sheetnames)
stats = wb['Pilots']

wiki = open('WR/MasterPilot.txt', 'w', encoding='utf-8')

# first row is titles
# ID NAME ROBOT ABILITY VALUE COST DESC
# All are str except ID (int)
# counting starts from 1

def writeln(*args, sep=' ', end='\n'):
    args = [str(i) for i in args]
    text = ' '.join(args)
    wiki.write(text+end)
    
def generate():
    pass