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
        time.sleep(0.15)
        pyautogui.click()
        print(f"Closed Popup")

def check_bag(screenshot, bag_image, bag_icon_image, threshold):
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
            time.sleep(0.15)
            pyautogui.click()
            print(f"Open Bag")

def check_crystal(game_window, screenshot, crtstal_image, threshold):
    template = cv2.imread(crtstal_image)

    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)

    if len(loc[0]) > 0:
        print(f"{game_window.title} Already Added")
        return True
    else:
        return False

def find_image_and_drag(main_template_path, target_template_path, crystal_images, popup_images, bag_image, bag_icon_image):
    threshold = 0.8

    while True:
        for game_window in gw.getWindowsWithTitle("FARM"):
            print(f"Switch to {game_window.title}")
            # Activate the game window
            game_window.activate()
            time.sleep(0.15)

            # Capture the game screen
            screenshot = np.array(ImageGrab.grab())
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

            check_bag(screenshot, bag_image, bag_icon_image, threshold)

            for image in popup_images:
                check_popup(screenshot, image, threshold)

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
                        print(f"{game_window.title} Not Added.")
                        # Calculate the center position of the found main image
                        main_center_x = main_loc[1][0] + main_w // 2
                        main_center_y = main_loc[0][0] + main_h // 2

                        # Drag object from main image to target image
                        print(f"Adding.")
                        pyautogui.moveTo(main_center_x, main_center_y, duration=0.1)
                        time.sleep(0.15)
                        pyautogui.mouseDown()
                        time.sleep(0.15)
                        pyautogui.moveTo(target_center_x, target_center_y + 10, duration=0.1)
                        time.sleep(0.15)
                        pyautogui.click(clicks=2, interval=0.2)
                        time.sleep(0.15)
                        print(f"Done.")
                        # Reset pointer
                        pyautogui.moveTo(960, 540, duration=0.1)
                        time.sleep(0.15)
                    else:
                        noti.send_line_notification(f"{game_window.title} Out of Crystal.")
                        substring_to_remove = "FARM"
                        new_window_title = game_window.title.replace(substring_to_remove,"")
                        set_window_title(game_window.title, new_window_title)
                        print(f"{game_window.title} Out of Crystal.")

            else:
                print(f"Skip {game_window.title}")

if __name__ == "__main__":
    main_path = "images/"
    main_template_path = f"{main_path}main_template_image.png"
    target_template_path = f"{main_path}target_template_image.png"
    bag_image = f"{main_path}bag.png"
    bag_icon_image = f"{main_path}bag_icon.png"
    crystal_images = [f"{main_path}low_crystal.png", f"{main_path}high_crystal.png"]
    popup_images = [f"{main_path}cancel.png", f"{main_path}close.png"]

    version = "v1.0.0"

    ctypes.windll.kernel32.SetConsoleTitleW(f"YG Crystal {version}")

    try:
        find_image_and_drag(main_template_path, target_template_path, crystal_images, popup_images, bag_image, bag_icon_image)
    except Exception as e:
        print(f"Something went wrong: {str(e)}")
