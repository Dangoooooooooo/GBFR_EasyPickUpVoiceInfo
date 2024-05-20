import tkinter as tk
import threading
from tkinter.scrolledtext import ScrolledText
import sys
from A2M1 import Action2Motion_Process
from M2E1 import Motion2Event_Process
from E2wID1 import Event2wemID_Process
from AltWavName1 import AltWavName_Process
from bnkRepack1 import bnkRepack_Process
from MixUp11 import MixUp_Process

def add_to_text_box(message):
    text_box.insert(tk.END, message + '\n')
    text_box.see(tk.END)

def execute_command(button, light, process_function):
    light.config(bg='yellow')
    button_text = button['text']  # 获取按钮上的文本
    add_to_text_box(f'=====================================================Command {button_text} is running.....')
    try:
        process_function()
        light.config(bg='green')
        add_to_text_box(f'=====================================================Command {button_text} completed successfully.')
    except Exception as e:
        light.config(bg='red')
        add_to_text_box(f'=====================================================An error occurred with command {button_text}: {e}')
    finally:
        button['state'] = 'normal'


# 更新命令列表为新的函数名
commands = [Action2Motion_Process, Motion2Event_Process, Event2wemID_Process, MixUp_Process, AltWavName_Process, bnkRepack_Process]

button_texts = ['(#1) PickUp [Action] <> [Motion]',
                '(#2) PickUp [Motion] <> [EventName]',
                '(#3) PickUp [EventName] <> [.wem]IDs',
                '(#4) Mix Them All',
                '(#5) Alter [.wav]s\' name to [EventName]s',
                '(#6) Repack [.bnk]s with new [.wem]s']

indicates = [
r'''PREPARE↓↓↓
①Copy <data\system\player\data\plxxxx\plxxxx_action.msg> to <.\Yours> folder.''',
r'''PREPARE↓↓↓
①Copy <data\pl\plxx00\> folder to <.\Yours> folder.''',
r'''PREPARE↓↓↓
①Finish #2.
②Copy <data\sound\(Japanese|English(US))\vo_plxx00(??).(bnk|pck)> to <.\Yours> folder.''',
r'''PREPARE↓↓↓
①Finish #1,#2,#3. 
②Suggest to install "GBFR Logs" and  copy <..\GBFR Logs\lang\(Your Language)\ui.json> to <.\Yours> folder.
···Recommended to only deal with [.bnk] files.
···If absolutely necessary, you can use "RingingBloom" unpacking [.pck] files to <.\Yours\wem_org\(FileName.pck)>.
>>DOWNLOAD "GBFR Logs"(Damage Meter Tools) on <github.com/false-spring/gbfr-logs>.
>>DOWNLOAD "RingingBloom" on <github.com/Silvris/RingingBloom>.''',
r'''PREPARE↓↓↓
①Fill up the sheet "YourRecord" in "GBFR#0Plxx00{Name}VoiceInfo_MixUp.xlsx". 
···MUST OPEN & EDIT & SAVE & CLOSE!!!''',
r'''PREPARE↓↓↓
①Finish #5. 
②Put your converted [.wem] files from "Wwise Launcher" to <.\Yours\wem_FromWwise\(FileName.bnk)>.
···If done, get your repacked [.bnk] files in <.\Yours\wem_FromWwise\(FileName.bnk_TEMP)\export\>.
>>DOWNLOAD <Wwise Launcher" on <audiokinetic.com>. ''']

# 创建主窗口
root = tk.Tk()
# root.configure(bg='light blue')
root.title(r"Granblue Fantasy:Relink - Easy PickUp Characters' VoiceInfo【BY:Dangoooooo\Bilibili  QQ:1041271418】")

First_text = r'''FIRSTLY, use "GBFRDataTools" unpacking "data.i" in game root directory <.\Steam\steamapps\common\Granblue Fantasy Relink>.
···You will get "data" folder in the same directory.
···Read the "PREPARE" instructions on the right. 
···Once finish preparation, press the button on the left.
>>DOWNLOAD "GBFRDataTools" on <github.com/Nenkai/GBFRDataTools>.'''
text_label = tk.Label(root, text=First_text, font=('Helvetica', 10, 'bold'), bg='pink', justify=tk.LEFT)
text_label.pack(fill='x')

for i in range(len(commands)):
    frame = tk.Frame(root)
    frame.pack(fill='x', padx=3, pady=3)

    button_text = button_texts[i]
    button = tk.Button(frame, text=f'Command {i+1}', width=33, height=1, font=('Helvetica', 10, 'bold'), anchor='w')
    button.pack(side='left')

    light = tk.Label(frame, width=1, height=1, bg='grey')
    light.pack(side='left')

    indicate = tk.Label(frame, text=indicates[i], width=96, fg='#0000FF', anchor='w', justify=tk.LEFT)
    indicate.pack(side='left', fill='x', padx=3)

    def make_callback(button, light, command):
        def callback():
            button['state'] = 'disabled'
            threading.Thread(target=execute_command, args=(button, light, command)).start()
        return callback

    button.config(command=make_callback(button, light, commands[i]),text=str(button_text))

text_box = ScrolledText(root, height=10)
text_box.pack(fill='both', expand=True)

class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)

    def flush(self):
        pass

# 重定向标准输出和错误到文本框
sys.stdout = TextRedirector(text_box)
sys.stderr = TextRedirector(text_box)

root.mainloop()
