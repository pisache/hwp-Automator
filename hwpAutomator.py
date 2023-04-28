import win32com.client as win32
import pandas as pd
import shutil
import random
import re

# user input
#xlFile = input("엑셀파일명을 입력 해주세요: \n")
#xlSheet = input("\n시트이름을 입력 해주세요: \n")
#maxNum = int(input("\n최대 문장 수를 입력 해주세요: \n"))

xlFile = "source.xlsx"
xlSheet = "sheet1"
maxNum = 757

'''
    Select next word
    Move selection left by 1 to remove space after selection
    apply underline
    cancel selection
'''
def underline():
    hwp.HAction.Run("MoveSelNextWord")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("CharShapeUnderline")
    hwp.HAction.Run("Cancel")
#end

cursor = 0

# Excel Extraction
numList = []
wordList = []
sentenceList = []
df = pd.read_excel(xlFile, xlSheet, dtype=str)
words = pd.read_excel(xlFile, xlSheet, na_values=['NA'], usecols="A", dtype=str)
senteces = pd.read_excel(xlFile, xlSheet, na_values=['NA'], usecols="B", dtype=str)
df2 = df


# list of 20 non-repeating random numbers
for i in range(100):
    x = random.randint(0, maxNum)
    while x in numList:
        x = random.randint(0, maxNum)
    while (x % 2) != 0:
        x = random.randint(0, maxNum)
    numList.append(x)
#end 

""" # Remove used words from excel and save it as new file
for i in range(len(numList)):
    df2 = df2.drop([numList[i], numList[i]+1])
#end
"""

words = words.astype(str)
for key, value in words.items():
    for i in range(len(numList)):
        wordList.append(value[numList[i]])
#end

for key, value in senteces.items():
    for i in range(len(numList)):
        sentenceList.append(value[numList[i]])
#end

# hangeul modification
infor = []

dic = {
    'q1': sentenceList[0], 'q2': sentenceList[1], 'q3': sentenceList[2], 'q4': sentenceList[3], 'q5': sentenceList[4],
    'q6': sentenceList[5], 'q7': sentenceList[6], 'q8': sentenceList[7], 'q9': sentenceList[8], 'q10': sentenceList[9],
    'q11': sentenceList[10], 'q12': sentenceList[11], 'q13': sentenceList[12], 'q14': sentenceList[13], 'q15': sentenceList[14],
    'q16': sentenceList[15], 'q17': sentenceList[16], 'q18': sentenceList[17], 'q19': sentenceList[18],  'q20': sentenceList[19],
    'q21': sentenceList[20],  'q22': sentenceList[21], 'q23': sentenceList[22], 'q24': sentenceList[23], 'q25': sentenceList[24],
    'q26': sentenceList[25], 'q27': sentenceList[26], 'q28': sentenceList[27], 'q29': sentenceList[28], 'q30': sentenceList[29],
    'q31': sentenceList[30], 'q32': sentenceList[31], 'q33': sentenceList[32], 'q34': sentenceList[33], 'q35': sentenceList[34],
    'q36': sentenceList[35], 'q37': sentenceList[36], 'q38': sentenceList[37], 'q39': sentenceList[38], 'q40': sentenceList[39],
    'q41': sentenceList[40], 'q42': sentenceList[41], 'q43': sentenceList[42], 'q44': sentenceList[43], 'q45': sentenceList[44],
    'q46': sentenceList[45], 'q47': sentenceList[46], 'q48': sentenceList[47], 'q49': sentenceList[48], 'q50': sentenceList[49],
    'q51': sentenceList[50], 'q52': sentenceList[51], 'q53': sentenceList[52], 'q54': sentenceList[53], 'q55': sentenceList[54],
    'q56': sentenceList[55], 'q57': sentenceList[56], 'q58': sentenceList[57], 'q59': sentenceList[58], 'q60': sentenceList[59],
    'q61': sentenceList[60], 'q62': sentenceList[61], 'q63': sentenceList[62], 'q64': sentenceList[63], 'q65': sentenceList[64],
    'q66': sentenceList[65], 'q67': sentenceList[66], 'q68': sentenceList[67], 'q69': sentenceList[68], 'q70': sentenceList[69],
    'q71': sentenceList[70], 'q72': sentenceList[71], 'q73': sentenceList[72], 'q74': sentenceList[73], 'q75': sentenceList[74],
    'q76': sentenceList[75], 'q77': sentenceList[76], 'q78': sentenceList[77], 'q79': sentenceList[78], 'q80': sentenceList[79],
    'q81': sentenceList[80], 'q82': sentenceList[81], 'q83': sentenceList[82], 'q84': sentenceList[83], 'q85': sentenceList[84],
    'q86': sentenceList[85], 'q87': sentenceList[86], 'q88': sentenceList[87], 'q89': sentenceList[88], 'q90': sentenceList[89],
    'q91': sentenceList[90], 'q92': sentenceList[91], 'q93': sentenceList[92], 'q94': sentenceList[93], 'q95': sentenceList[94],
    'q96': sentenceList[95], 'q97': sentenceList[96], 'q98': sentenceList[97], 'q99': sentenceList[98], 'q100': sentenceList[99]
}
infor.append(dic)

shutil.copyfile(r"./test.hwp",r"./test_out.hwp")

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

# Path for production
#hwp.Open(r"C:\dist\test_out.hwp")

# Path for testing
hwp.Open(r"D:\HoJun\dev\hwpAutomator\test_out.hwp")
fieldList = [i for i in hwp.GetFieldList().split("\x02")] 

for field in fieldList:
    hwp.PutFieldText(f'{field}{{{{0}}}}', infor[0][field])
hwp.MovePos(2)

# set the cursor on the first sentece.
hwp.InitScan(option=0x04, Range=0x0007)
for i in range(4):
    hwp.GetText()
    hwp.MovePos(201)
#end
hwp.ReleaseScan()
hwp.HAction.Run("MoveSelLineEnd")
hwp.InitScan(option=0x02, Range=0x00ff)

hwp.HAction.Run("MoveLineBegin")
hwp.HAction.Run("MoveDown")
hwp.HAction.Run("MoveNextWord")
hwp.HAction.Run("MoveSelLineEnd")

""" hwp.InitScan(option=0x02, Range=0x00ff)
id, scanString = hwp.GetText()
for m in re.finditer(wordList[1], scanString):
        position = m.start()
hwp.HAction.Run("Cancel")
hwp.MovePos(1, 3, position + 10)
underline()
print(scanString)
print(wordList[1]) """

# Loop to underline:
for i in range(len(wordList)):
    id, scanString = hwp.GetText()
    print(scanString)
    print(wordList[i])
    for m in re.finditer(wordList[i], scanString):
        position = m.start()
    hwp.HAction.Run("Cancel")
    hwp.MovePos(1, cursor+2, position + 10)
    underline()

    hwp.HAction.Run("MoveLineBegin")
    if i % 5 == 4:
        for j in range(4):
            hwp.HAction.Run("MoveDown")
        cursor += 3
    #end
    else:
        hwp.HAction.Run("MoveDown")
        cursor += 1
    #end
    hwp.HAction.Run("MoveNextWord")
    hwp.ReleaseScan()
    hwp.HAction.Run("MoveSelLineEnd")
    if len(scanString) > 98:
        hwp.HAction.Run("MoveSelDown")
        cursor += 1
    #end
    hwp.InitScan(option=0x02, Range=0x00ff)
#end

hwp.HAction.Run("Cancel")
