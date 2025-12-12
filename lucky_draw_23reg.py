import os
import random
import tkinter as tk
from tkinter import BooleanVar
from PIL import Image, ImageTk
from tkvideo import tkvideo
import pandas as pd
import threading
import time

# -------------------- GLOBAL SETUP --------------------
root = tk.Tk()
root.title("Christmas Dinner")
root.overrideredirect(True)

w = root.winfo_screenwidth()
h = root.winfo_screenheight()
root.geometry(f"{w}x{h}")
root.configure(bg="white")

osc_stage_shown = False
first_prize_label = None
first_prize_at_plane = False   # tracks position
first_prize_animating = False  # prevents double trigger

# -------------------- GLOBAL COLOR THEME (Normal Draw Only) --------------------
# Choose ONE theme by setting THEME = "gold_red", "caramel", or "green_gold"

THEME = ""   # <-- change here to try different styles

if THEME == "gold_red":
    # Warm gold + burgundy
    NORMAL_NUM_FG = "#FFD66B"      # warm gold
    NORMAL_NAME_FG = "#FFF3E0"     # creamy white
    NORMAL_BG = "#7A1A1A"          # warm burgundy label background

elif THEME == "caramel":
    # Warm soft caramel + chocolate
    NORMAL_NUM_FG = "#F4C16F"      # caramel gold
    NORMAL_NAME_FG = "#FFF6E8"     # cream
    NORMAL_BG = "#8C5E3C"          # chocolate brown

elif THEME == "green_gold":
    # Evergreen warm gold
    NORMAL_NUM_FG = "#FFDE72"      # soft bright gold
    NORMAL_NAME_FG = "#FFF9E6"     # warm white
    NORMAL_BG = "#2F6A3B"          # soft evergreen green

else:
    # Fallback to your original cold theme
    NORMAL_NUM_FG = "#b0872e"
    NORMAL_NAME_FG = "#0A2A43"
    NORMAL_BG = "#FFFFFF"


# -------------------- SAFE LABEL WRAPPER --------------------
class SafeLabel(tk.Label):
    """Label wrapper that silently ignores config errors after destruction"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._is_alive = True

    def config(self, **kwargs):
        if self._is_alive:
            try:
                super().config(**kwargs)
            except tk.TclError:
                # Widget was destroyed, silently ignore
                self._is_alive = False

    def configure(self, **kwargs):
        self.config(**kwargs)

    def destroy(self):
        self._is_alive = False
        try:
            super().destroy()
        except:
            pass


# Load main background
bg_img = Image.open("2025bg.jpg").resize((w, h))
bg_photo = ImageTk.PhotoImage(bg_img)
background_label = tk.Label(root, image=bg_photo)
background_label.place(x=0, y=0)

# Intro sequence files
content_files = ["3.jpg", "2.jpg", "1.mp4"]
current_index = -1
draw_started = False
current_widget = None
video_player = None
video_thread = None
video_playing = False

# Prize setup
prize_number = 75
count_num = 0
previous_winner_label = None
previous_winner_label_number = None

# Load Excel
data = pd.read_excel("2025_list.xlsx").values.tolist()
guest_number = len(data)

# -------------------- SPACE WAIT VARIABLE --------------------
space_pressed = BooleanVar(value=False)


def wait_for_space():
    space_pressed.set(False)
    root.bind("<space>", lambda e: space_pressed.set(True))
    root.wait_variable(space_pressed)
    root.bind("<space>", handle_space)  # restore main handler

# -------------------- VIDEO CLEANUP --------------------
def cleanup_video():
    """Force cleanup of video resources"""
    global video_player, video_playing, current_widget, current_index

    video_playing = False
    video_player = None

    # destroy widget
    if current_widget:
        try:
            current_widget.destroy()
        except:
            pass
        current_widget = None

# -------------------- INTRO DISPLAY --------------------
def show_intro_content():
    global current_index, current_widget, video_player, video_playing

    if draw_started:
        return

    # Clean up previous content
    cleanup_video()

    current_index += 1

    # INTRO FINISHED → start draw
    if current_index >= len(content_files):
        start_first_draw()
        return

    content = content_files[current_index]

    if content.endswith(".jpg"):
        # Clear old widgets
        for widget in root.winfo_children():
            if widget != background_label:
                try:
                    widget.destroy()
                except:
                    pass

        img = Image.open(content).resize((w, h))
        photo = ImageTk.PhotoImage(img)
        lbl = tk.Label(root, image=photo)
        lbl.image = photo
        lbl.place(x=0, y=0)
        current_widget = lbl

    elif content.endswith(".mp4"):
        # Clear old widgets
        for widget in root.winfo_children():
            if widget != background_label:
                try:
                    widget.destroy()
                except:
                    pass

        # Use SafeLabel for video to prevent thread errors
        video_lbl = SafeLabel(root, bg="black", width=w, height=h)
        video_lbl.place(x=0, y=0, width=w, height=h)
        current_widget = video_lbl

        # Force update before playing video
        root.update_idletasks()

        try:
            # Create and play video
            player = tkvideo(content, video_lbl, loop=0, size=(w, h))
            video_player = player
            video_playing = True

            # Play in a controlled way
            player.play()

        except Exception as e:
            print(f"Error playing video: {e}")
            # If video fails, move to next content
            show_intro_content()


# -------------------- DESTROY INTRO / START DRAW --------------------
def start_first_draw():
    global draw_started, current_widget, video_player, background_label, video_playing

    draw_started = True
    video_playing = False

    # Cleanup video
    cleanup_video()

    # Clear all widgets
    for widget in root.winfo_children():
        try:
            widget.destroy()
        except:
            pass

    # Recreate main background
    background_label = tk.Label(root, image=bg_photo)
    background_label.place(x=0, y=0)

    # Ensure keyboard focus after intro
    root.focus_force()
    root.focus_set()
    root.update()

    # DO NOT CALL perform_draw() HERE
    # The first draw should only happen when user hits ENTER

# -------------------- MAIN DRAW FUNCTION --------------------
def perform_draw():
    global previous_winner_label, previous_winner_label_number
    global count_num, draw_started, current_widget, video_player, video_playing

    # Load drawn history
    if os.path.exists("history_log.txt"):
        used_numbers = [line.strip() for line in open("history_log.txt")]
    else:
        used_numbers = []

    # Remaining prizes
    if os.path.exists("Winner_List.txt"):
        drawn_list = open("Winner_List.txt").readlines()
    else:
        drawn_list = []

    prize_left = prize_number - len(drawn_list)

    # --------------------------------------------
    # 🎄 ALL PRIZES DRAWN
    # --------------------------------------------
    if prize_left <= 0:
        lbl = tk.Label(root, text="Merry Christmas!",
                       font=("Arial", 50), fg="gold", bg="#152238")
        lbl.place(relx=0.5, rely=0.2, anchor="center")
        return

    # --------------------------------------------
    # ⭐ OSC SEQUENCE FOR LAST 5
    # --------------------------------------------
    global osc_stage_shown

    # When prize_left is 3,2,1 → show the image/video ONLY ONCE
    if prize_left == 3 and not osc_stage_shown:
        osc_stage_shown = True
        show_top_image("3.jpg")
        return

    if prize_left == 2 and not osc_stage_shown:
        osc_stage_shown = True
        show_top_image("2.jpg")
        return

    if prize_left == 1 and not osc_stage_shown:
        osc_stage_shown = True
        show_top_video("1.mp4")
        return

    # --------------------------------------------
    # If we reach here and prize_left <= 5,
    # it means we ALREADY showed the image/video,
    # and now ENTER was pressed again → continue draw
    # --------------------------------------------

    # Pick winner
    while True:
        idx = random.randint(0, guest_number - 1)
        if str(idx) not in used_numbers:
            break

    emp_no = data[idx][0]
    first = str(data[idx][1]).strip()
    last = str(data[idx][2]).strip()
    last_short = last.strip()[0].upper() + "."
    result = f"{first} {last_short} ({emp_no})"

    # Save logs
    with open("history_log.txt", "a") as f:
        f.write(str(idx) + "\n")
    with open("Winner_List.txt", "a") as f:
        f.write(f"#{prize_left}: {result}\n")

    # --------------------------------------------
    # ⭐ DISPLAY TOP-4 RESULT
    # --------------------------------------------
    # --------------------------------------------
    # ⭐ DISPLAY TOP-4 RESULT (with custom colors)
    # --------------------------------------------
    if prize_left <= 4:

        # When #4 appears → clear EVERYTHING except background
        if prize_left == 4:
            for widget in root.winfo_children():
                if widget != background_label:
                    try:
                        widget.destroy()
                    except:
                        pass

            previous_winner_label = None
            previous_winner_label_number = None

        # Destroy previous top labels if they exist
        if previous_winner_label:
            previous_winner_label.destroy()
        if previous_winner_label_number:
            previous_winner_label_number.destroy()

        # -------------------------------
        # 🎨 COLOR LOGIC FOR TOP 5
        # -------------------------------
        if prize_left == 4:
            # First two of top 5 → same as before
            num_fg = "#b0872e"
            name_fg = "#0A2A43"
            bg = "#FFFFFF"

        elif prize_left in (3, 2):
            # Keep current style (same as above actually)
            num_fg = "#b0872e"
            name_fg = "#0A2A43"
            bg = "#FFFFFF"

        elif prize_left == 1:
            display_1st_prize_name(f"#1: {result}")
            return

        # -------------------------------
        # POSITION FOR TOP 5
        # -------------------------------

        # Default center position
        pos_x = 0.35
        pos_y = 0.55

        # Raise #3 and #2 slightly higher (OSC request)
        if prize_left in (3, 2):
            pos_y = 0.30   # <-- lifted upward (was 0.55)

        num_lbl = tk.Label(root, text=f"#{prize_left}: ",
                           font=("Arial", 60, "bold"), fg=num_fg, bg=bg)
        num_lbl.place(relx=pos_x, rely=pos_y, anchor="se")
        previous_winner_label_number = num_lbl

        name_lbl = tk.Label(root, text=result,
                            font=("Arial", 60, "bold"), fg=name_fg, bg=bg)
        name_lbl.place(relx=pos_x, rely=pos_y, anchor="sw")
        previous_winner_label = name_lbl

        osc_stage_shown = False
        return

    # --------------------------------------------
    # NORMAL GRID MODE FOR PRIZES > 5
    # --------------------------------------------

    # When 20 entries (4×5) are shown → clear screen
    if count_num >= 20:
        for widget in root.winfo_children():
            if widget != background_label:
                try:
                    widget.destroy()
                except:
                    pass
        count_num = 0

    # --------------------------------------------
    # SPECIAL LAYOUT FOR 4TH DRAW (11 prizes)
    # --------------------------------------------
    if prize_left <= 11 and prize_left > 4:
        # 3 columns:
        # col 0 = 4 rows
        # col 1 = 4 rows
        # col 2 = 3 rows

        if count_num < 4:
            col = 0
            row = count_num
        elif count_num < 8:
            col = 1
            row = count_num - 4
        else:
            col = 2
            row = count_num - 8

    else:
        # DEFAULT BEHAVIOR (unchanged)
        col = count_num // 5
        row = count_num % 5

    # --------------------------------------------
    # DIFFERENT GRID CENTERING FOR #16 → #6
    # --------------------------------------------
    if prize_left <= 11 and prize_left > 4:
        # Centered 3-column layout for 4th draw
        base_x = 0.15 + 0.25 * col
        base_y = 0.32 + 0.12 * row
    elif 5 < prize_left <= 15:
        base_x = 0.15 + 0.35 * col
        base_y = 0.32 + 0.12 * row
    else:
        # Original placement for all other normal-grid draws
        base_x = 0.08 + 0.21 * col
        base_y = 0.32 + 0.12 * row

    # Smaller font
    grid_font_size = 20

    num_lbl = tk.Label(
        root,
        text=f"#{prize_left}: ",
        font=("Arial", grid_font_size, "bold"),
        fg=NORMAL_NUM_FG,
        bg=NORMAL_BG
    )
    num_lbl.place(relx=base_x, rely=base_y, anchor="se")

    name_lbl = tk.Label(
        root,
        text=result,
        font=("Arial", grid_font_size),
        fg=NORMAL_NAME_FG,
        bg=NORMAL_BG
    )
    name_lbl.place(relx=base_x, rely=base_y, anchor="sw")

    count_num += 1

def show_top_image(filename):
    global current_widget

    # Clear previous widgets
    for widget in root.winfo_children():
        if widget != background_label:
            try:
                widget.destroy()
            except:
                pass

    img = Image.open(filename).resize((w, h))
    photo = ImageTk.PhotoImage(img)
    lbl = tk.Label(root, image=photo)
    lbl.image = photo
    lbl.place(x=0, y=0)

    # Bind ENTER ONCE ONLY:
    root.bind("<Return>", handle_after_top_stage)

def show_top_video(filename):
    global current_widget

    for widget in root.winfo_children():
        if widget != background_label:
            try:
                widget.destroy()
            except:
                pass

    video_lbl = SafeLabel(root, bg="black", width=w, height=h)
    video_lbl.place(x=0, y=0, width=w, height=h)

    player = tkvideo(filename, video_lbl, loop=0, size=(w, h))
    player.play()

    # Bind ENTER ONCE ONLY:
    root.bind("<Return>", handle_after_top_stage)

def handle_after_top_stage(event):
    # Restore default ENTER behavior AFTER drawing
    perform_draw()
    root.bind("<Return>", handle_enter)

# -------------------- KEY HANDLERS --------------------
def handle_enter(event):
    global draw_started

    # BEFORE DRAW STARTS → handle intro
    if not draw_started:
        # If video is playing, kill it first
        if video_playing:
            cleanup_video()
            root.after(100, show_intro_content)
            return
        else:
            show_intro_content()
        return

    # AFTER DRAW STARTS → normal draw
    perform_draw()

def handle_space(event):
    global first_prize_label, first_prize_at_plane, first_prize_animating

    # During intro, space does nothing
    if not draw_started:
        return

    # Only allow SPACE control for #1 animation
    if first_prize_label is None:
        perform_draw()
        return

    # Prevent spam while animation is running
    if first_prize_animating:
        return

    # Toggle destination
    if not first_prize_at_plane:
        # move to plane position
        animate_first_prize_move(
            0.5, 0.5,      # start
            0.5, 0.66      # end
        )
        first_prize_at_plane = True
    else:
        # move back to original
        animate_first_prize_move(
            0.5, 0.66,     # start
            0.5, 0.5       # end
        )
        first_prize_at_plane = False


def handle_redraw(event):
    global osc_stage_shown
    global first_prize_label, first_prize_at_plane, first_prize_animating

    if not os.path.exists("Winner_List.txt"):
        return

    # Remove last winner
    lines = open("Winner_List.txt").readlines()
    if not lines:
        return

    with open("Winner_List.txt", "w") as f:
        f.writelines(lines[:-1])

    # Remove last entry from history_log
    if os.path.exists("history_log.txt"):
        hist_lines = open("history_log.txt").readlines()
        if hist_lines:
            with open("history_log.txt", "w") as f:
                f.writelines(hist_lines[:-1])

    # ---------------------------------------------------
    # FIX 1 — Prevent showing OSC stages again
    # ---------------------------------------------------
    osc_stage_shown = True

    # ---------------------------------------------------
    # FIX 2 — If currently showing #1 prize, remove label
    # ---------------------------------------------------
    if first_prize_label is not None:
        try:
            first_prize_label.destroy()
        except:
            pass
        first_prize_label = None
        first_prize_at_plane = False
        first_prize_animating = False

    # ---------------------------------------------------
    # Immediately draw again
    # ---------------------------------------------------
    perform_draw()

def display_1st_prize_name(winner_text):
    global first_prize_label, first_prize_at_plane

    first_prize_at_plane = False  # start centered

    first_prize_label = tk.Label(
        root,
        text=winner_text,
        font=("Arial", 50, "bold"),
        fg="gold",
        bg="#152238",
        padx=10,
        pady=10
    )
    first_prize_label.place(relx=0.5, rely=0.5, anchor="center")

def animate_first_prize_move(start_x, start_y, end_x, end_y):
    global first_prize_label, first_prize_animating

    first_prize_animating = True

    steps = 120
    duration = 3000   # 3 seconds
    delay = duration // steps

    dx = (end_x - start_x) / steps
    dy = (end_y - start_y) / steps

    # Font resize logic
    initial_size = int(first_prize_label.cget("font").split()[1])
    target_size = 9 if end_y > start_y else 50
    df = (initial_size - target_size) / steps

    def do_step(step=0):
        global first_prize_animating

        if step > steps:
            first_prize_animating = False
            return

        new_x = start_x + dx * step
        new_y = start_y + dy * step
        first_prize_label.place(relx=new_x, rely=new_y, anchor="center")

        new_font = int(initial_size - df * step)
        first_prize_label.config(font=("Arial", new_font, "bold"))

        root.after(delay, lambda: do_step(step + 1))

    do_step(0)

def shrink_and_move_label(label, start_x, start_y, end_x, end_y,
                          duration, target_font_size,
                          return_after_ms=20000):
    """
    Smooth movement + shrink animation, then return after a delay.
    """
    steps = 250
    delay = duration / steps

    # movement delta per step
    dx = (end_x - start_x) / steps
    dy = (end_y - start_y) / steps

    # extract initial font size
    initial_font_size = int(label.cget("font").split()[1])
    df = (initial_font_size - target_font_size) / steps

    def animate_forward(step=0):
        if step <= steps:
            # move
            label.place(relx=start_x + dx * step,
                        rely=start_y + dy * step,
                        anchor="center")

            # shrink
            new_size = max(target_font_size,
                           int(initial_font_size - df * step))
            label.config(font=("Arial", new_size, "bold"))

            root.after(int(delay), lambda: animate_forward(step + 1))

        else:
            # After reaching destination → stay, then return
            root.after(return_after_ms, animate_return)

    def animate_return(step=0):
        if step <= steps:
            label.place(relx=end_x - dx * step,
                        rely=end_y - dy * step,
                        anchor="center")

            new_size = min(initial_font_size,
                           int(target_font_size + df * step))
            label.config(font=("Arial", new_size, "bold"))

            root.after(int(delay), lambda: animate_return(step + 1))

    animate_forward()

def handle_escape(event):
    """Exit application"""
    cleanup_video()
    root.quit()
    root.destroy()

def skip_video(event):
    """Skip video if it's playing"""
    global video_playing
    if video_playing and not draw_started:
        cleanup_video()
        root.after(100, show_intro_content)


# -------------------- KEY BINDINGS --------------------
root.bind("<Return>", handle_enter)  # ENTER = intro
root.bind("<Shift_R>", handle_redraw)  # SHIFT RIGHT = redo
root.bind("<Escape>", handle_escape)  # ESC = exit
root.bind("s", skip_video)  # S = skip video
root.bind("<space>", handle_space)

# -------------------- RUN --------------------
try:
    root.mainloop()
except KeyboardInterrupt:
    cleanup_video()
    root.destroy()