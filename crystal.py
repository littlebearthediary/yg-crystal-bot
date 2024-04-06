import cv2
import numpy as np
from PIL import ImageGrab
import pygetwindow as gw
import pyautogui
import time
import noti
import win32gui
import ctypes

def set_window_title(old_title, new_title):
    target_window_handle = win32gui.FindWindow(None, old_title)
    win32gui.SetWindowText(target_window_handle, new_title)

def check_popup(screenshot, template_image, threshold):
    # Load the template image
    template = cv2.imread(template_image)
    h, w, _ = template.shape

    # Match template
    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)

    if len(loc[0]) > 0:
        center_x = loc[1][0] + w // 2
        center_y = loc[0][0] + h // 2

        pyautogui.moveTo(center_x, center_y, duration=0.1)
        time.sleep(0.1)
        pyautogui.click()
        time.sleep(0.1)
        pyautogui.moveTo(1270, 540, duration=0.1)
        print(f"Closed popup.")
        return True
    return False

def check_connection(game_window, screenshot, warning_image, threshold):
    template = cv2.imread(warning_image)

    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)

    if len(loc[0]) > 0:
        substring_to_remove = "xONLINE"
        new_window_title = game_window.title.replace(substring_to_remove, "")
        set_window_title(game_window.title, new_window_title)
        
        print(f"{game_window.title} Disconnected.")
        noti.send_line_notification(f"{game_window.title}\nDisconnected.")
        return False
    return True

def check_bag_open(screenshot, bag_image, bag_icon_image, threshold):
    # Load the template image
    image = cv2.imread(bag_image)
    bag_icon = cv2.imread(bag_icon_image)
    bag_icon_h, bag_icon_w, _ = bag_icon.shape

    # Match template
    res = cv2.matchTemplate(screenshot, image, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)

    bag_res = cv2.matchTemplate(screenshot, bag_icon, cv2.TM_CCOEFF_NORMED)
    bag_loc = np.where(bag_res >= threshold)

    if not len(loc[0]) > 0:
        if len(bag_loc[0] > 0):
            bag_icon_x = bag_loc[1][0] + bag_icon_w // 2
            bag_icon_y = bag_loc[0][0] + bag_icon_h // 2

            pyautogui.moveTo(bag_icon_x, bag_icon_y, duration=0.1)
            time.sleep(0.1)
            pyautogui.click()
            print(f"Open bag.")
            return True
        pyautogui.moveTo(930, 750, duration=0.1)
        time.sleep(0.1)
        pyautogui.click()
        print(f"Open bag.")
        return True
    return False

def check_crystal(game_window, screenshot, crtstal_image, threshold):
    template = cv2.imread(crtstal_image)

    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)

    if len(loc[0]) > 0:
        print(f"{game_window.title} Already added.")
        return True
    else:
        return False
    
def check_empty(screenshot, empty_image, threshold):
    template = cv2.imread(empty_image)

    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)

    if len(loc[0]) > 0:
        return True
    else:
        return False
    
def check_game_panel_open(screenshot, panel_images, threshold):
    for image in panel_images:
        template = cv2.imread(image)

        res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        loc = np.where(res >= threshold)

        if len(loc[0]) > 0:
            return True
    return False

def find_image_and_drag(main_template_path, target_template_path, crystal_images, popup_images, bag_image, bag_icon_image, empty_image, panel_images, warning_image):
    threshold = 0.8

    for game_window in gw.getWindowsWithTitle("FARM"):
        if not ("xONLINE" in game_window.title):
            new_window_title = game_window.title + "xONLINE"
            set_window_title(game_window.title, new_window_title)
        if ("xFULL" in game_window.title):
            substring_to_remove = "xFULL"
            new_window_title = game_window.title.replace(substring_to_remove, "")
            set_window_title(game_window.title, new_window_title)

    while True:
        for game_window in gw.getWindowsWithTitle("xONLINE"):
            print(f"Switch to {game_window.title}")
            # Activate the game window
            game_window.activate()
            time.sleep(0.1)

            # Capture the game screen
            screenshot = np.array(ImageGrab.grab())
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

            # if not check_connection(game_window, screenshot, warning_image, threshold):
            #     continue

            for image in popup_images:
                if check_popup(screenshot, image, threshold):
                    screenshot = np.array(ImageGrab.grab())
                    screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

            if check_game_panel_open(screenshot, panel_images, threshold):
                print(f"Wait for character to get back.")
                continue

            if check_bag_open(screenshot, bag_image, bag_icon_image, threshold):
                screenshot = np.array(ImageGrab.grab())
                screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

            if check_empty(screenshot,empty_image, threshold):
                crystal_check = []

                # Check if any crystal image is found
                for image in crystal_images:
                    if check_crystal(game_window, screenshot, image, threshold):
                        crystal_check.append(True)

                if len(crystal_check) == 0:
                    # Match target template
                    target_template = cv2.imread(target_template_path)
                    target_h, target_w, _ = target_template.shape
                    target_res = cv2.matchTemplate(screenshot, target_template, cv2.TM_CCOEFF_NORMED)
                    target_loc = np.where(target_res >= threshold)

                    if len(target_loc[0]) > 0:
                        # Calculate the center position of the found target image
                        target_center_x = target_loc[1][0] + target_w // 2
                        target_center_y = target_loc[0][0] + target_h // 2

                        # Match main template
                        main_template = cv2.imread(main_template_path)
                        main_h, main_w, _ = main_template.shape
                        main_res = cv2.matchTemplate(screenshot, main_template, cv2.TM_CCOEFF_NORMED)
                        main_loc = np.where(main_res >= threshold)

                        if len(main_loc[0]) > 0:
                            print(f"{game_window.title} Not added.")
                            # Calculate the center position of the found main image
                            main_center_x = main_loc[1][0] + main_w // 2
                            main_center_y = main_loc[0][0] + main_h // 2

                            # Drag object from main image to target image
                            print(f"Adding.")
                            pyautogui.moveTo(main_center_x, main_center_y, duration=0.1)
                            time.sleep(0.1)
                            pyautogui.mouseDown()
                            time.sleep(0.1)
                            pyautogui.moveTo(target_center_x, target_center_y + 30, duration=0.1)
                            time.sleep(0.1)
                            pyautogui.click(clicks=2, interval=0.2)
                            time.sleep(0.1)
                            print(f"Done.")
                            # Reset pointer
                            pyautogui.moveTo(1270, 540, duration=0.1)
                        else:
                            substring_to_remove = "xONLINE"
                            new_window_title = game_window.title.replace(substring_to_remove,"")
                            set_window_title(game_window.title, new_window_title)
                            noti.send_line_notification(f"{game_window.title}\nOut of crystal.")
                            print(f"{game_window.title} Out of crystal.")
                else:
                    print(f"Skip {game_window.title}")
            else:
                if not ("xFULL" in game_window.title):
                    new_window_title = game_window.title + "xFULL"
                    set_window_title(game_window.title, new_window_title)
                    noti.send_line_notification(f"{game_window.title}\nBag is full.")
                    print(f"{game_window.title} Bag is full.")

if __name__ == "__main__":
    main_path = "images/"
    main_template_path = f"{main_path}main_template_image.png"
    target_template_path = f"{main_path}target_template_image.png"
    bag_image = f"{main_path}bag.png"
    bag_icon_image = f"{main_path}bag_icon.png"
    crystal_images = [f"{main_path}low_crystal.png", f"{main_path}high_crystal.png"]
    popup_images = [f"{main_path}cancel.png", f"{main_path}close.png", f"{main_path}popup_message.png"]
    empty_image = f"{main_path}empty.png"
    panel_images = [f"{main_path}storage_panel.png", f"{main_path}shop_panel.png", f"{main_path}npc_panel.png"]
    warning_image = f"{main_path}warning.png"

    version = "v1.2.3"

    ctypes.windll.kernel32.SetConsoleTitleW(f"YG Crystal {version}")

    find_image_and_drag(main_template_path, target_template_path, crystal_images, popup_images, bag_image, bag_icon_image, empty_image, panel_images, warning_image)