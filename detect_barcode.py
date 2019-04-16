import helper
import argparse
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from openpyxl import Workbook
import tkinter as tk
import os
import time
from openpyxl import Workbook

lst = []
times = []
excelfile = Workbook()
row = excelfile.active
row.append(['รหัสประจำตัว', 'เวลา'])

def scan():
    camera = cv2.VideoCapture(0)
    while True:
        (grabbed, frame) = camera.read()

        if not grabbed:
            break

        box = helper.detect(frame)
        mask = np.zeros_like(frame)
        try:
            cv2.drawContours(frame, [box], -1, 255, -1)
            cv2.drawContours(mask, [box], -1, 255, -1)
        except:
            print('Please move QRcode')
            continue

        # crop image
        out = np.zeros_like(frame)
        out[mask == 255] = frame[mask == 255]

        (x, y) = np.where(mask == 255)[:2]
        (topx, topy) = (np.min(x), np.min(y))
        (bottomx, bottomy) = (np.max(x), np.max(y))
        out = out[topx:bottomx+1, topy:bottomy+1]

        x = decode(out)
        if len(x) != 0:
            v = x[0].data.decode('UTF-8')
            print('Found IT!!')
            if v not in lst:
                lst.append(v)
                times.append(time.asctime( time.localtime(time.time()) ))

        cv2.imshow("Frame", frame)

        key = cv2.waitKey(1) & 0xFF

        if key == ord("q"):
            break
    for v, t in zip(lst, times):
        row.append((v, t))
        id = ('รหัสนักศึกษา :' + str(idl) + '  ' + time.asctime( time.localtime(time.time()) ) + '\n')
        txt.insert(0.0, id)
    camera.release()
    cv2.destroyAllWindows()

def add(abc=None):

    idl=ent.get()
    print(idl)
    row.append([idl, time.asctime( time.localtime(time.time()) )])
    id = ('รหัสนักศึกษา :' + str(idl) + '  ' + time.asctime( time.localtime(time.time()) ) + '\n')
    txt.insert(0.0,id)
    ent.delete(0, tk.END)


def save(abc=None):
    pop_up = tk.Tk()

    pop_up.title('save')
    label = tk.Label(pop_up,text='ชื่อไฟล์ :',font='time24').grid(row=0)
    pop_up_ent = tk.Entry(pop_up, width=45, bd=5)
    pop_up_ent.grid(row=0,column=1)
    pop_up_but = tk.Button(pop_up, text="Save",font='time24',bg="green" \
                           ,fg="white"\
                           ,width = 25,height = 2 \
                           ,bd = 5)
    pop_up_but.grid(row=1,columnspan=2)

    def save_ja(abc=None):
        filename = pop_up_ent.get()
        filename = filename + '.xlsx'
        excelfile.save(filename)
        print('Excel Saved :', filename)
        print('SAVED!')
        pop_up.destroy()
        gui.destroy()
    pop_up.bind('<Return>', save_ja)

gui=tk.Tk()
gui.bind('<Return>', add)
gui.geometry('455x320')
gui.title('SCANNER DEMO 1.0')

ll=tk.Label(gui,text='รหัสนักศึกษา :',font='time24')
ll.grid(row=0)


ent = tk.Entry(gui,width=45,bd=5)
ent.grid(row=0,column=1)

but=tk.Button(gui,text="Scan",font='time24',bg="green",fg="white" \
              ,command = scan,width = 25,height = 2,bd = 5)

but.grid(row=1,columnspan=2)

a = tk.Button(gui, text ="SAVE",font='time24',bg="gold",fg="black"\
              ,command = save,width = 25,height = 2,bd = 5)

a.grid(row=2,columnspan=2)


txt=tk.Text(gui,width=55,height=10,bd=5)
txt.grid(row=3,columnspan=2)
gui.mainloop()




