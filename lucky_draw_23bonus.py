import os
import random
import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
import pandas as pd

# ---------------------------------------------------
# WINDOW SETUP
# ---------------------------------------------------
root = tk.Tk()
root.title("Bonus Draw")
root.overrideredirect(True)

w = root.winfo_screenwidth()
h = root.winfo_screenheight()
root.geometry(f"{w}x{h}")
root.configure(bg="white")

# -----------------------------
# FONT SETTINGS (Unified)
# -----------------------------
BONUS_FONT_SIZE = 23
BONUS_FONT = ("Arial", BONUS_FONT_SIZE)
BONUS_FONT_BOLD = ("Arial", BONUS_FONT_SIZE, "bold")

# ---------------------------------------------------
# LOAD BACKGROUND
# ---------------------------------------------------
bg_img = Image.open("2025bg.jpg").resize((w, h))
bg_photo = ImageTk.PhotoImage(bg_img)

background_label = tk.Label(root, image=bg_photo)
background_label.place(x=0, y=0)

# ---------------------------------------------------
# LOAD EXCEL
# ---------------------------------------------------
data = pd.read_excel("2025_list.xlsx").values.tolist()
guest_number = len(data)

# ---------------------------------------------------
# BONUS DRAW SETTINGS
# ---------------------------------------------------
count_num = 0                # grid counter
BONUS_LIMIT = 999            # unlimited unless you want to cap it
BONUS_FILE = "Bonus_List.txt"
HISTORY_FILE = "Bonus_history.txt"

# COLORS (exact same as your main regular draw)
NORMAL_NUM_FG = "#b0872e"
NORMAL_NAME_FG = "#0A2A43"
NORMAL_BG = "#FFFFFF"

# ---------------------------------------------------
# CLEAN GRID (after 15)
# ---------------------------------------------------
def clear_grid():
    for wdg in root.winfo_children():
        if wdg != background_label:
            wdg.destroy()

# ---------------------------------------------------
# MAIN BONUS DRAW FUNCTION
# ---------------------------------------------------
def perform_bonus_draw(event=None):
    global count_num

    # Load used indices
    if os.path.exists(HISTORY_FILE):
        used = [line.strip() for line in open(HISTORY_FILE)]
    else:
        used = []

    # Prevent infinite when all drawn
    if len(used) >= guest_number:
        lbl = tk.Label(root, text="All guest numbers used!",
                       font=("Arial", 40, "bold"),
                       fg="gold", bg="#152238")
        lbl.place(relx=0.5, rely=0.1, anchor="center")
        return

    # Random unique number
    while True:
        idx = random.randint(0, guest_number - 1)
        if str(idx) not in used:
            break

    # Extract guest info
    emp_no = data[idx][0]
    first = str(data[idx][1]).strip()
    last = str(data[idx][2]).strip()
    last_short = last.strip()[0].upper() + "."
    result_text = f"{first} {last_short} ({emp_no})"

    # Save logs
    with open(HISTORY_FILE, "a") as f:
        f.write(str(idx) + "\n")

    with open(BONUS_FILE, "a") as f:
        f.write(f"Bonus: {result_text}\n")

    # ------------------------------
    # CLEAR SCREEN AFTER 15 ENTRIES
    # ------------------------------
    if count_num >= 15:
        clear_grid()
        count_num = 0

    # ------------------------------
    # PLACE IN 5×3 GRID
    # ------------------------------
    col = count_num // 5
    row = count_num % 5

    base_x = 0.10 + 0.26 * col  # adjust as needed
    base_y = 0.30 + 0.12 * row  # adjust as needed

    # Draw labels
    num_lbl = tk.Label(root,
                       text="Bonus:",
                       font=BONUS_FONT,
                       fg=NORMAL_NUM_FG,
                       bg=NORMAL_BG)
    num_lbl.place(relx=base_x, rely=base_y, anchor="se")

    name_lbl = tk.Label(root,
                        text=result_text,
                        font=BONUS_FONT,
                        fg=NORMAL_NAME_FG,
                        bg=NORMAL_BG)
    name_lbl.place(relx=base_x, rely=base_y, anchor="sw")

    count_num += 1

# ---------------------------------------------------
# REDRAW LAST BONUS
# ---------------------------------------------------
def undo_bonus(event=None):
    global count_num

    if not os.path.exists(BONUS_FILE):
        return
    if not os.path.exists(HISTORY_FILE):
        return

    # remove last line from BONUS file
    lines = open(BONUS_FILE).readlines()
    if lines:
        with open(BONUS_FILE, "w") as f:
            f.writelines(lines[:-1])

    # remove last line from history
    hist = open(HISTORY_FILE).readlines()
    if hist:
        with open(HISTORY_FILE, "w") as f:
            f.writelines(hist[:-1])

    # rebuild grid
    clear_grid()
    count_num = 0

    # redraw all from BONUS file
    for line in open(BONUS_FILE).readlines():
        # parse line: "Bonus: NAME"
        try:
            name = line.split("Bonus: ")[1].strip()
            # simulate draw without saving
            draw_bonus_label(name)
        except:
            pass

def draw_bonus_label(name):
    """
    Helper used when redrawing grid from file.
    """
    global count_num

    if count_num >= 15:
        clear_grid()
        count_num = 0

    col = count_num // 5
    row = count_num % 5

    base_x = 0.10 + 0.26 * col
    base_y = 0.36 + 0.12 * row

    num_lbl = tk.Label(root,
                       text="Bonus:",
                       font=BONUS_FONT,
                       fg=NORMAL_NUM_FG,
                       bg=NORMAL_BG)
    num_lbl.place(relx=base_x, rely=base_y, anchor="se")

    name_lbl = tk.Label(root,
                        text=name,
                        font=BONUS_FONT,
                        fg=NORMAL_NAME_FG,
                        bg=NORMAL_BG)
    name_lbl.place(relx=base_x, rely=base_y, anchor="sw")

    count_num += 1

# ---------------------------------------------------
# EXIT HANDLER
# ---------------------------------------------------
def quit_app(event=None):
    root.destroy()

# ---------------------------------------------------
# KEY BINDINGS
# ---------------------------------------------------
root.bind("<Return>", perform_bonus_draw)
root.bind("<Shift_R>", undo_bonus)
root.bind("<Escape>", quit_app)

# ---------------------------------------------------
# RUN PROGRAM
# ---------------------------------------------------
root.mainloop()
