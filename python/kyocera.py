import shutil
import os
import logging
import time
import pyautogui
import psutil
import platform
import subprocess
import sys
from datetime import datetime
import webbrowser

PRINTER_NAME_FILE_PORT = ""
PRINTER_IP_REAL = ""
DIRECTORY_PATH = ""

"""
Description: 
    This setups basic functionality of script like logging and global variable assignment from user
    Asks user to input specific data required for the proper runtime of script
"""
def setup():
    global PRINTER_IP_REAL, PRINTER_NAME_FILE_PORT, DIRECTORY_PATH

    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    log_file = os.path.join(exe_dir, datetime.now().strftime("%m-%d-%Y_%H-%M_runtime.log"))

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    logging.info("Runtime started.")
    logging.warning("This is a warning message.")
    logging.error("This is an error message.")

    print("-------------------------")
    print("Commands")
    print("--update")
    print("--readme")
    print("--run")
    print("")
    choice = input("Please enter a command: ")

    if choice.strip().lower() == "--update":
        project_page = "https://github.com/zachreid-96/Document_Box_Transfer_Tool"
        webbrowser.open(project_page)
        input("Press any key to exit...")
    elif choice.strip().lower() == "--readme":
        readme_path = extract_README()
        os.startfile(readme_path)
        input("Press any key to exit...")
    elif choice.strip().lower() == "--run":

        print("PLEASE follow these strict instructions to setup the environment correctly...")
        readme_path = extract_README()
        os.startfile(readme_path)

        exit_condition = input("Is the environment setup correctly? (Yes/No): ")

        if not exit_condition.lower().startswith('y'):
            logging.error("Failed to setup environment correctly.")
            input("Press any key to exit...")

        PRINTER_NAME_FILE_PORT = input("Printer name of FILE Port printer: ")
        PRINTER_IP_REAL = input("IP of the device the documents need sent to: ")
        DIRECTORY_PATH = input("Directory of the documents: ")

    else:
        logging.error("Failed to parse command.")
        input("Press any key to exit...")

    return


def get_running_procs():
    return {proc.pid: proc.name().lower() for proc in psutil.process_iter()}


def detect_new_proc(old_procs, timeout=5):
    start_time = time.time()
    while time.time() - start_time < timeout:
        new_proc = get_running_procs()
        detected = {pid: name for pid, name in new_proc.items() if pid not in old_procs}

        if detected:
            print(max(detected, key=detected.get))
            return max(detected, key=detected.get)

        time.sleep(0.5)

    return None


def force_close_proc(pid):
    try:
        proc = psutil.Process(pid)
        proc.terminate()
        time.sleep(2)

        if proc.is_running():
            proc.kill()

    except psutil.NoSuchProcess:
        pass


def create_PRN_files(directory):
    excluded_extensions = [".prn"]

    file_list = []
    for root, _, files in os.walk(directory):
        for file in files:
            if not any(file.endswith(ext.lower()) for ext in excluded_extensions):
                file_list.append(os.path.join(root, file))

    for file in file_list:

        path_parts = file.split("\\")
        save_directory = "\\".join(path_parts[0:-1])
        save_filename = path_parts[-1].split('.')[0] + ".prn"

        control_proc = get_running_procs()

        try:
            os.startfile(file, "print")
            time.sleep(2)

            detected_application = detect_new_proc(control_proc)

            pyautogui.write(save_filename)
            time.sleep(1)

            pyautogui.hotkey("alt", "d")
            time.sleep(1)

            pyautogui.write(save_directory)
            pyautogui.press("enter")
            time.sleep(1)

            pyautogui.hotkey("alt", "s")
            time.sleep(1)

            force_close_proc(detected_application)

        except Exception as e:
            logging.error(f"Failed to get submit job: {e}")

    return


def inject_PJL_commands(directory):
    file_list = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".prn"):
                file_list.append(os.path.join(root, file))

    for file in file_list:
        path_parts = file.split("\\")
        box_num = str(path_parts[-2]).zfill(4)
        save_filename = path_parts[-1].split('.')[0]
        with open(file, 'rb') as f, open("temp.prn", "wb") as temp_file:
            while True:
                line = f.readline()
                if line.startswith(b"@PJL SET ECONOMODE="):
                    temp_file.write(line)
                    temp_file.write(b'@PJL SET HOLD=KUSERBOX\n')
                    temp_file.write(f'@PJL SET KUSERBOXID="{box_num}"\n'.encode())
                    temp_file.write(b'@PJL SET KUSERBOXPASSWORD=""\n')
                elif line.startswith(b'@PJL SET JOBNAME='):
                    temp_file.write(f'@PJL SET JOBNAME="{save_filename}"\n'.encode())
                    break
                else:
                    temp_file.write(line)

            shutil.copyfileobj(f, temp_file)

        shutil.copy("./temp.prn", file)
        send_LPR_file(file)
    return


def send_LPR_file(file):
    system = platform.system().lower()
    if system == "windows":
        try:
            subprocess.call(['where', 'lpr'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError:
            logging.error("LPR NOT ENABLED. Please enable LPR on workstation.")
    else:
        try:
            subprocess.call(['which', 'lpr'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError:
            logging.error("LPR NOT INSTALLED. Please install LPR on workstation.")

    subprocess.call(['lpr', '-S', PRINTER_IP_REAL, '-P', '9100', file],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def extract_README():

    readme_path = os.path.join(os.path.dirname(sys.executable), "README.md")

    if getattr(sys, 'frozen', False):
        bundled_readme = os.path.join(sys._MEIPASS, "README.md")
        if not os.path.exists(readme_path):
            shutil.copy(bundled_readme, readme_path)
            logging.info("Extracted README.md.")
    else:
        logging.warning("Failed to extract README.md.")

    return readme_path


setup()
create_PRN_files(DIRECTORY_PATH)
inject_PJL_commands(DIRECTORY_PATH)
