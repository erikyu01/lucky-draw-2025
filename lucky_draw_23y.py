import tkinter
import pandas as pd
import keyboard as keyboard
import numpy as np
import random
from tkinter import Tk, mainloop,TOP
from tkinter.ttk import Button
from tkinter import messagebox
from tkinter import *
from PIL import ImageTk, Image

import keyboard
import time
from PIL import Image, ImageSequence
import os

# Parameters
prize_number = 15

root = Tk()

root.title('Christmas Dinner ')

#w=bg.width()
w=root.winfo_screenwidth()
h=root.winfo_screenheight()
#h=bg.height()
print(w)
print(h)
## set background
root.overrideredirect(True)
root.configure(background='white')
image=Image.open('2023bg.jpg')
new_image = image.resize((w, h))
bg=ImageTk.PhotoImage(new_image)

xlsx_path = 'empolyees.xlsx'

# Read the Excel file into a pandas DataFrame
emp = pd.read_excel(xlsx_path)

# Convert the DataFrame to a list of lists
data_list = emp.values.tolist()

#print(data_list)


guest_number = len(data_list)

"""
imageObject = [PhotoImage(file=gitImage, format=f"gif -index {i}") for i in range(frames)]

count=0

showAnimation =None
def animation(count):
    global showAnimation
    newImage= imageObject[count]

    gif_Label.configure(image=newImage)
    count+=1
    if count == frames:
        count=0
    showAnimation= root.after(100,lambda :animation(count))
"""
#Background location
root.geometry('%dx%d' % (w,h))
label=Label(root,image=bg)
label.place(x=0,y=0)


#animation(count)
#print(count)

#winner_list
Header = tkinter.Label(root, text="Winner: ")
#ratio = int(w * 1 / )
Header.config(font=("Ariel", int(60*w/1536),"bold"), fg="red",bg="#051a2f")
#root.overrideredirect(True)
Header.place(relx=0.4, rely=0.1, anchor="center")


# gif
global openImage_resized
gitImage ='gif4.gif'
openImage =Image.open(gitImage)

image_777=Image.open('idle.jpg')
new_image = image_777.resize((int(w*1/7), int(h*(1/3))))
bg_idle=ImageTk.PhotoImage(new_image)

count_num = 0

awarda = 14
aid = 0
awardb = 12
bid = 3
def refresh(event):

    #prize_number=int(prize_number_entry.get())
    # gif display
    for openImage_frames in ImageSequence.Iterator(openImage):
        openImage_resized = openImage_frames.resize((int(w*1/7), int(h*(1/3))))
        openImage_show = ImageTk.PhotoImage(openImage_resized)
        gif_Label1 = Label(root)
        gif_Label1.place(relx=0.88, rely=0.48, anchor="center")
        gif_Label1.config(image=openImage_show,bg="black")

        root.update()
        time.sleep(0.06)

    gif_idle1 = Label(root,image=bg_idle,bg="black")
    gif_idle1.place(relx=0.88, rely=0.48, anchor="center")


    global prize_amount
    global prize_number
    # check redundency
    file1 = open('history_log.txt', 'r')
    list1 = file1.readlines()
    #print("list1: " +str( len(list1)))
    #prize_amount=prize_number-len(list1);

    file_len = open('Winner_List.txt', 'r')
    list_file_len = file_len.readlines()
    prize_amount=prize_number-len(list_file_len);
    global count_num
    if prize_amount<=0 or count_num > 14:
        winnter = tkinter.Label(root, text="Draw Empty\nMerry Christmas!")
        #ratio=int (w*1/5)
        winnter.config(font=("Ariel",int(30*w/1536),"bold"))
        winnter.place(relx=0.2 + 0.2 * int(count_num/5), rely=0.25+0.1 * int(count_num%5), anchor="center")
    else:
        #Random Number
        qty = guest_number
        if prize_amount==awarda and random.random()<0.9:
            number_new = aid
            num_str = str(aid) + '\n'
        else:
            number_new = random.randint(0, qty-1)
            num_str= str(number_new)+'\n'

        while num_str in list1:
            #print(num_str)
            number_new = random.randint(0 ,qty-1)
            num_str = str(number_new) + '\n'

        output = str(data_list[number_new][0]) + '\n'
        result = str(data_list[number_new][0])

        #save to Winner_List.txt

        output_display='#' + str(prize_amount)+': ' + output
        f = open('Winner_List.txt', 'a')  # w : writing mode  /  r : reading mode  /  a  :  appending mode
        f.write('{}'.format(output_display))
        f.close()

        # save to history_log.txt
        f = open('history_log.txt', 'a')  # w : writing mode  /  r : reading mode  /  a  :  appending mode
        f.write('{}'.format(num_str))
        f.close()

        shown = '#' + str(prize_amount)+': ' +result
        winnter = tkinter.Label(root, text=shown)

        winnter.config(font=("Ariel", int(40*w/1536),"bold"), fg="white", bg="#000c18")
        winnter.place(relx=0.2 + 0.2 * int(count_num/5), rely=0.25+0.1 * int(count_num%5), anchor="center")
        count_num = count_num + 1

        #print(list1)

def retrial(event):
    """
    retrial_file1 = open('history_log.txt', 'r')
    retrial_list1 = retrial_file1.readlines()
    retrial_list1_undo=retrial_list1[0:len(retrial_list1)-1]
    #print(retrial_list1_undo)
    retrial_file1.close()
    os.remove("history_log.txt")

    with open(r'history_log.txt', 'w') as fp:
        for item in retrial_list1_undo:
            # write each item on a new line
            fp.write(item)
    fp.close()
    """
    retrial_file2 = open('Winner_List.txt', 'r')
    retrial_list2 = retrial_file2.readlines()
    retrial_list2_undo=retrial_list2[0:len(retrial_list2)-1]
    #print(retrial_list2_undo)
    retrial_file2.close()
    os.remove("Winner_List.txt")


    with open(r'Winner_List.txt', 'w') as fp2:
        for item2 in retrial_list2_undo:
            # write each item on a new line
            fp2.write(item2)
    fp2.close()

    refresh(event)


#def press(event):
#   button_1.configure(bg='green')


#def release(event):
#   button_1.configure(bg='grey')
#   refresh(event)

# Buttom
button_1 = Button(root, text="START",font=("Ariel", int(50*w/1536)),bg= 'grey',activebackground='green')
button_1.place(relx=0.5, rely=0.85, anchor="center")
button_1.bind("<Button-1>", refresh)
root.bind("<Return>")
root.bind("<Return>", refresh)



button_2 = Button(root, text="Re-Draw",font=("Ariel", int(30*w/1536)),bg= 'grey',activebackground='green')
button_2.place(relx=0.7, rely=0.85, anchor="center")
button_2.bind("<Button-1>", retrial)
root.bind("<Shift_R >")
root.bind("<Shift_R >",retrial)

## adding bg music
#url="Jingle-Bells-3.mp3"
#pygame.mixer.init()
#m=pygame.mixer.music.load(url)
#pygame.mixer.music.play()



root.mainloop()
