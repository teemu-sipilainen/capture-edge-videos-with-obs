import configparser
import csv
import os
import pyautogui
import pygetwindow as gw
import re
import time
from obswebsocket import obsws, requests
from pywinauto.application import Application
from screeninfo import get_monitors
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

def load_config(filename, section):
    config = configparser.ConfigParser()
    config.read(filename)
    return config[section]

CONFIG_FILE = 'config.ini'

paths_config = load_config(CONFIG_FILE, "PATHS")
OBS_PATH = paths_config['OBS_PATH']
CMD_PATH = paths_config['CMD_PATH']
NEW_WORKING_DIRECTORY = paths_config['NEW_WORKING_DIRECTORY']

videos = []

with open('videos-to-capture.csv', newline='') as csvfile:
    csvreader = csv.DictReader(csvfile)
    for row in csvreader:
        record = row['record'] == 'True'
        video_name = row['link']
        section = int(row['hours'])
        start_time = int(row['minutes'])
        end_time = int(row['seconds'])
        
        videos.append([record, video_name, section, start_time, end_time])



def video_length_in_seconds(hours, minutes, seconds): 
    return hours * 60 * 60 + minutes * 60 + seconds


######################
##### OBS STUDIO #####
######################

def connect_to_obs_web_socket_server(host, port, password):
    obs_web_socket_server = obsws(host, port, password)
    obs_web_socket_server.connect()
    return obs_web_socket_server

def get_file_name_from_video_link(video_link):

    with open('captured-video-file-name-pattern.txt', 'r') as file:
        pattern = file.read().strip()
    
    match = re.search(pattern, video_link)
    
    if match:
        return match.group(1)
    else:
        return None
    
def set_obs_web_socket_recording_information(obs_web_socket_server, folder, filename):
    try:
        response = obs_web_socket_server.call(
            requests.SetProfileParameter(
                parameterCategory="SimpleOutput",
                parameterName="FilePath",
                parameterValue=folder
            )
        )

        if response.status == 'success':
            print(f"Recording directory set: {folder}")
        else:
            print(f"Recording directory set failed: {response}")

        response = obs_web_socket_server.call(
            requests.SetProfileParameter(
                parameterCategory="Output",
                parameterName="FilenameFormatting",
                parameterValue=filename
            )
        )

        if response.status == 'success':
            print(f"Recording filename set: {filename}")
        else:
            print(f"Recording filename set failed: {response}")
    
    except Exception as e:
        print(f"Error: {e}")
        print(f"Exception: {type(e).__name__}, {str(e)}")

def set_obs_web_socket_recording_filename(obs_web_socket_server, filename):
    try:
        response = obs_web_socket_server.call(
            requests.SetProfileParameter(
                parameterCategory="Output",
                parameterName="FilenameFormatting",
                parameterValue=filename
            )
        )

        if response.status == 'success':
            print(f"Recording filename set: {filename}")
        else:
            print(f"Recording filename set failed: {response}")
    
    except Exception as e:
        print(f"Error: {e}")
        print(f"Exception: {type(e).__name__}, {str(e)}")

def set_obs_web_socket_recording_filepath(obs_web_socket_server, filepath):
    try:
        response = obs_web_socket_server.call(
            requests.SetProfileParameter(
                parameterCategory="SimpleOutput",
                parameterName="FilePath",
                parameterValue=filepath
            )
        )

        if response.status == 'success':
            print(f"Recording directory set: {filepath}")
        else:
            print(f"Recording directory set failed: {response}")
    
    except Exception as e:
        print(f"Error: {e}")
        print(f"Exception: {type(e).__name__}, {str(e)}")


def start_obs_web_socket_recording(obs_web_socket_server):
    try:
        response = obs_web_socket_server.call(requests.StartRecord())
        print("Response:", response)
        if response.status == 'success':
            print("Start recording success.")
        else:
            print("Start recording failed:", response)
    except Exception as e:
        print(f"Error: {e}")
        print(f"Exception: {type(e).__name__}, {str(e)}")

def stop_obs_web_socket_recording(obs_web_socket_server):
    try:
        response = obs_web_socket_server.call(requests.StopRecord())
        print("Response:", response)
        if response.status == 'success':
            print("Stop recording success.")
        else:
            print("Stop recording failed:", response)
    except Exception as e:
        print(f"Error: {e}")
        print(f"Exception: {type(e).__name__}, {str(e)}")

def disconnect_from_obs_web_socket_server(obs_web_socket_server):
    obs_web_socket_server.disconnect()
    print("WebSocket connection closed.")

def change_working_directory(new_working_directory):
    try:
        os.chdir(new_working_directory)
        print(f"Current working directory changed to: {os.getcwd()}")
    except FileNotFoundError:
        print(f"Directory '{new_working_directory}' not found.")

def set_obs_studio_active(obs_studio_app): 
    obs_studio_app.set_focus()

def start_obs_studio_recording(obs_studio_app):

    set_obs_studio_active(obs_studio_app)
    
    time.sleep(5)

    obs_studio_app.type_keys("{F5}")

def start_obs_studio(): 
    
    change_working_directory(NEW_WORKING_DIRECTORY)
    os.startfile(OBS_PATH)
    
    time.sleep(5) 

    return get_obs_studio_window()

def get_obs_studio_window(): 

    obs_studio_app = None

    while not obs_studio_app:
        try:
            obs_studio_app = gw.getWindowsWithTitle('OBS 30.0.2')[0]
            print("OBS Studio application window was found:", obs_studio_app)
            
        except IndexError:
            time.sleep(1)
            print("OBS Studio application window was not found")
    
    app = Application().connect(handle=obs_studio_app._hWnd)

    obs_studio_window = app.window(handle=obs_studio_app._hWnd)
    
    return obs_studio_window

def stop_obs_studio_recording(obs_studio_app):
    set_obs_studio_active(obs_studio_app)
    obs_studio_app.type_keys("{F6}")

def close_obs_studio(obs_studio_app): 
    set_obs_studio_active(obs_studio_app)

###################
##### BROWSER #####
###################

def set_browser_active(browser_app): 

    browser_app.set_focus()

def get_browser_window(): 

    browser_window = None
    while not browser_window:
        try:
            browser_window = gw.getWindowsWithTitle('Profiili 1 – Microsoft\u200b Edge')[0]
            print(browser_window)
            print("Browser application window was found")
            return browser_window
            
        except IndexError:
            time.sleep(1)
            return None
    
    return None

def get_browser_app(browser_window): 

    app = Application().connect(handle=browser_window._hWnd)

    browser_app = app.window(handle=browser_window._hWnd)
    
    return browser_app


def browser_open(): 

    driver = webdriver.Edge()  # webdriver.Chrome() or webdriver.Edge()
    time.sleep(1)
    
    driver.maximize_window()
    time.sleep(2)

    return driver

def browser_goto(driver, browser_app, link): 
    set_browser_active(browser_app)
    driver.get(link)
    time.sleep(5)

def browser_close(driver, browser_app): 
    set_browser_active(browser_app)
    driver.quit()

def move_mouse_away_from_video(driver, browser_app): 
    set_browser_active(browser_app)
    actions = ActionChains(driver)
    actions.move_by_offset(200, 0).perform() 

def move_mouse_to_monitor(monitor): 
    if (monitor): 
        pyautogui.moveTo( monitor.x + monitor.width / 2, monitor.y + monitor.height / 2 )
    
def get_mouse_monitor(): 
    monitors = get_monitors()
    if len(monitors) > 1:
        current_mouse_x, current_mouse_y = pyautogui.position()
        print(current_mouse_x)
        print(current_mouse_y)
        
        current_monitor = None
        for monitor in monitors:
            if (monitor.x <= current_mouse_x < monitor.x + monitor.width) and (monitor.y <= current_mouse_y < monitor.y + monitor.height):
                return monitor
            
    return None

def get_browser_monitor(browser_window): 
    if (browser_window):
        monitors = get_monitors()
        if len(monitors) > 1:

            for monitor in monitors:
                print("Monitor:", monitor)
                if (monitor.x <= browser_window.left * -1 < monitor.x + monitor.width and monitor.y <= browser_window.top * -1 < monitor.y + monitor.height):
                    print("Browser is in the monitor:", monitor)
                    return monitor
    
    return None

def get_not_browser_monitor(browser_monitor): 
    if (browser_monitor):
        monitors = get_monitors()
        if len(monitors) > 1:

            for monitor in monitors:
                if (monitor == browser_monitor):
                    continue
                else: 
                    return monitor
                
    return None

def move_mouse():
    current_mouse_x, current_mouse_y = pyautogui.position()
    pyautogui.moveTo(current_mouse_x + 200, current_mouse_y)

def move_mouse_to_video(driver, browser_app):
    set_browser_active(browser_app)
    video_area = get_video_area(driver, browser_app)
    actions = ActionChains(driver)
    actions.move_to_element(video_area).perform()

def move_mouse_to_video_and_doubleclick(driver, browser_app): 
    set_browser_active(browser_app)
    video_area = get_video_area(driver, browser_app)
    actions = ActionChains(driver)
    actions.move_to_element(video_area).perform()
    actions.click(video_area).perform()
    time.sleep(0.2)
    actions.click(video_area).perform()

def get_video_area(driver, browser_app): 
    set_browser_active(browser_app)
    video_area = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "player-region-container"))
    )
    return video_area

def prepare_video(driver, browser_app):

    set_browser_active(browser_app)

    move_mouse_to_video_and_doubleclick(driver, browser_app)
    
    time.sleep(1)

    full_screen_with_shortcuts(driver, browser_app)
    time.sleep(1)


    for i in range(0,3):
        rewind_with_shortcuts3(driver, browser_app)
        time.sleep(1)

def full_screen_with_shortcuts(driver, browser_app):
    set_browser_active(browser_app)
    move_mouse_to_video(driver, browser_app)

    fullscreen_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Full screen']"))
    )
    fullscreen_button.click()

def rewind_with_shortcuts3(driver, browser_app): 
    set_browser_active(browser_app)
    move_mouse_to_video(driver, browser_app)
    rewind_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Skip 10 seconds backward']"))
    )
    rewind_button.click()

def rewind_with_shortcuts2(driver, browser_app): 
    set_browser_active(browser_app)
    move_mouse_to_video(driver, browser_app)
    seek_bar = driver.find_element(By.XPATH, "//input[@aria-label='Progress bar']")
    driver.execute_script("arguments[0].value = 0;", seek_bar)
    seek_bar.send_keys("\n")
    time.sleep(2)

# Videon toistaminen pikanäppäimien avulla
def rewind_with_shortcuts(driver, browser_app):
    set_browser_active(browser_app)
    move_mouse_to_video(driver, browser_app)
    try:
        actions = ActionChains(driver)

        actions.key_down(Keys.ALT)

        actions.send_keys('J')

        actions.key_up(Keys.ALT)

        for i in range(0, 2):
            actions.perform()

        time.sleep(2)

    except Exception as e:
        print(f"Exception in using the video: {e}")
        driver.quit()

def play_video_with_shortcuts(driver, browser_app):
    set_browser_active(browser_app)
    move_mouse_to_video(driver, browser_app)

    play_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Play']"))
    )
    play_button.click()

def stop_video_with_shortcuts(driver, browser_app):
    set_browser_active(browser_app)
    play_video_with_shortcuts(driver, browser_app)

def play_video(driver, browser_app):
    set_browser_active(browser_app)
    

    play_video_with_shortcuts(driver, browser_app)
    move_mouse_away_from_video(driver, browser_app)

##################################################

def main():

    obs_studio_app = get_obs_studio_window()

    driver = browser_open()

    time.sleep(5)

    browser_window = get_browser_window()
    print(browser_window)

    browser_app = get_browser_app(browser_window)
    print(browser_app)

    obs_websocket_config = load_config(CONFIG_FILE, 'OBS_WEBSOCKET')
    HOST = obs_websocket_config['HOST']
    PORT = int(obs_websocket_config['PORT'])
    PASSWORD = obs_websocket_config['PASSWORD']

    obs_web_socket_server = connect_to_obs_web_socket_server(HOST, PORT, PASSWORD)

    time.sleep(5)

    for video in videos:

        if (not video[0]):
            continue

        mouse_monitor = get_mouse_monitor()
        print("Mouse monitor:", mouse_monitor)

        browser_monitor = get_browser_monitor(browser_window)
        print("Browser monitor:", browser_monitor)

        not_browser_monitor = get_not_browser_monitor(browser_monitor)
        print("Other than browser monitor:", not_browser_monitor)
        
        video_link = video[1]
        video_length = video_length_in_seconds(video[2], video[3], video[4])

        file_name = get_file_name_from_video_link(video_link)
        set_obs_web_socket_recording_filename(obs_web_socket_server, file_name)

        browser_goto(driver, browser_app, video_link)

        move_mouse_to_monitor(browser_monitor)
        prepare_video(driver, browser_app)
        
        start_obs_web_socket_recording(obs_web_socket_server)

        time.sleep(5)

        play_video(driver, browser_app)
        move_mouse_to_monitor(not_browser_monitor)
        time.sleep(video_length+10)

        stop_obs_web_socket_recording(obs_web_socket_server)

        time.sleep(5)

    disconnect_from_obs_web_socket_server(obs_web_socket_server)

##################################################

if __name__ == "__main__":
    main()
