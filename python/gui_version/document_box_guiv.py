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
from pathlib import Path
from logger_setup import setup


def setup_backup():
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    log_file = os.path.join(exe_dir, datetime.now().strftime("%m-%d-%Y_%H-%M_runtime.log"))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file)
        ]
    )

    logging.info("Runtime started.")
    logging.warning("This is a warning message.")
    logging.error("This is an error message.")

    return logging.getLogger()


def get_running_procs():
    return {proc.pid: proc.name().lower() for proc in psutil.process_iter()}


def detect_new_proc(old_procs, timeout=5):
    start_time = time.time()
    while time.time() - start_time < timeout:
        new_proc = get_running_procs()
        detected = {pid: name for pid, name in new_proc.items() if pid not in old_procs}

        if detected:
            return max(detected, key=detected.get)

        time.sleep(0.5)

    return None


def force_close_proc(pid):
    try:
        proc = psutil.Process(pid)
        proc.terminate()
        time.sleep(3)

        if proc.is_running():
            proc.kill()

    except psutil.NoSuchProcess:
        pass

    return None


def create_PRN_files(file):
    logger = setup()
    path_parts = file.parts
    save_directory = f"{path_parts[0]}{'\\'.join(path_parts[1:-1])}"
    save_filename = path_parts[-1].split('.')[0] + ".prn"

    control_proc = get_running_procs()

    try:
        os.startfile(file, "print")
        time.sleep(4)

        detected_application = detect_new_proc(control_proc)

        pyautogui.write(save_filename)
        time.sleep(2)

        pyautogui.hotkey("alt", "d")
        time.sleep(2)

        pyautogui.write(save_directory)
        pyautogui.press("enter")
        time.sleep(2)

        pyautogui.hotkey("alt", "s")
        time.sleep(2)

        force_close_proc(detected_application)

    except Exception as e:
        logger.error(f"Failed to get submit job: {e}")

    return


def inject_PJL_commands(file, printer_ip):
    logger = setup()
    path_parts = file.parts
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
    send_LPR_file(file, printer_ip)
    return


def send_LPR_file(file, printer_ip):
    logger = setup()
    system = platform.system().lower()
    if system == "windows":
        try:
            subprocess.call(['where', 'lpr'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError:
            logger.error("LPR NOT ENABLED. Please enable LPR on workstation.")
    else:
        try:
            subprocess.call(['which', 'lpr'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError:
            logger.error("LPR NOT INSTALLED. Please install LPR on workstation.")

    subprocess.call(['lpr', '-S', printer_ip, '-P', '9100', file],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def extract_README():
    logger = setup()
    readme_path = os.path.join(os.path.dirname(sys.executable), "README.md")

    if getattr(sys, 'frozen', False):
        bundled_readme = os.path.join(sys._MEIPASS, "README.md")
        if not os.path.exists(readme_path):
            shutil.copy(bundled_readme, readme_path)
            logger.info("Extracted README.md.")
    else:
        logger.warning("Failed to extract README.md.")

    return readme_path
