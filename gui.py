import tkinter as tk
from tkinter import filedialog

def browse():
    filename = filedialog.askopenfilename()
    pathEnt.delete(0, tk.END)
    pathEnt.insert(0, filename)
    print("Selectedf file:", filename)

root = tk.Tk()
root.title("hwp Automator")
root.geometry("540x320")
root.resizable(False, False)

welcomeLabel = tk.Label(root, text="""hwp Automator for Harry's Education\n""",
justify='left')
welcomeLabel.place(x=1, y=1)

issueLabel = tk.Label(root, text="""knwon issues:
한글에서 한줄이 넘어가는 문장 사용시 그 다음 줄 단어 밑줄이 누락되는 현상\n""",
justify='left')
issueLabel.place(x=1, y=35)

warnLabel = tk.Label(root, text="""주의사항:
test.hwp파일을 수정하지 말아주세요.
해당 폴더를 C드라이브 이외에 장소에 보관 시 작동하지 않습니다.\n""",
justify='left')
warnLabel.place(x=1, y=80)

pathLabel = tk.Label(root, text="엑셀 파일 위치:")
pathLabel.place(x=1, y=150)

pathEnt = tk.Entry(root, width = 50)
pathEnt.place(x=5, y=170)

lookupBtn = tk.Button(root, text="찾아보기", command=browse)
lookupBtn.place(x=370, y=165)

sentence = tk.Text(root, width = 50, height= 5, background='gray91')
sentence.place(x=5, y=210)

runBtn = tk.Button(root, text="실행", height=2, width=6)
runBtn.place(x=370, y=210)

creditLabel = tk.Label(root, text = """\n제작자: 이호준
환경 테스팅: 송서영
python 3.8.10""",
justify='right')
creditLabel.place(x=420, y=250)

root.mainloop()