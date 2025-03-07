import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import win32print
import document_box_guiv as db
import ctypes
from pathlib import Path
import threading
from datetime import datetime
from logger_setup import setup

user32 = ctypes.windll.user32


def toggle_inputs(state):
    global user32
    if state == "disable":
        user32.BlockInput(True)
    elif state == "enable":
        user32.BlockInput(False)


class DocumentBoxGUI:
    def __init__(self, r):
        self.root = r
        self.root.title("Document Box Transfer Tool")
        self.root.geometry("550x250")

        self.file_list = []
        self.file_list_size = 0
        self.file_list_progress = 0

        self.prn_list = []
        self.prn_list_size = 0
        self.prn_list_progress = 0

        dir_frame = ttk.Frame(r)
        dir_frame.pack(pady=5, fill=tk.X, padx=10)

        self.dir_label = ttk.Label(dir_frame, text="Select Directory:")
        self.dir_label.pack(side=tk.LEFT)
        self.dir_entry = ttk.Entry(dir_frame, width=50)
        self.dir_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.dir_button = ttk.Button(dir_frame, text="Browse", command=self.select_directory)
        self.dir_button.pack(side=tk.RIGHT)

        file_printer_frame = ttk.Frame(r)
        file_printer_frame.pack(pady=5, fill=tk.X, padx=10)

        self.printer_label = ttk.Label(file_printer_frame, text="Select Printer:")
        self.printer_label.pack(side=tk.LEFT)
        self.printer_var = tk.StringVar()
        self.printer_dropdown = ttk.Combobox(file_printer_frame, textvariable=self.printer_var, state="readonly",
                                             width=47)
        self.printer_dropdown.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.refresh_button = ttk.Button(file_printer_frame, text="Refresh Printers", command=self.load_file_printers)
        self.refresh_button.pack(side=tk.RIGHT)

        self.load_file_printers()

        ip_printer_frame = ttk.Frame(r)
        ip_printer_frame.pack(pady=5, fill=tk.X, padx=10)

        self.ip_label = ttk.Label(ip_printer_frame, text="Enter Printer IP Address:")
        self.ip_label.pack(side=tk.LEFT)
        self.ip_entry = ttk.Entry(ip_printer_frame, width=50)
        self.ip_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.refresh_button = ttk.Button(ip_printer_frame, text="Ping Printer", command=self.ping_printer)
        self.refresh_button.pack(side=tk.RIGHT)

        self.start_button = ttk.Button(r, text="Start Process", command=self.start_process)
        self.start_button.pack(pady=5)

        progress_frame = ttk.Frame(r)
        progress_frame.pack(pady=5, fill=tk.Y, padx=10)

        self.ping_var = tk.StringVar(value="Ping... ")
        self.progress_var = tk.StringVar(value="Processing... ")
        self.file_list_var = tk.StringVar(value="Getting File List... ")
        self.processing_var = tk.StringVar(value="Creating .prn Files... ")
        self.injecting_var = tk.StringVar(value="Injecting PJL Commands... ")
        self.sending_var = tk.StringVar(value="Sending File via LPR... ")

        self.label_ping = tk.Label(progress_frame, textvariable=self.ping_var, anchor="w", width=40)
        self.label_progress = tk.Label(progress_frame, textvariable=self.progress_var, anchor="w", width=40)
        self.label_file_list = tk.Label(progress_frame, textvariable=self.file_list_var, anchor="w", width=40)
        self.label_processing = tk.Label(progress_frame, textvariable=self.processing_var, anchor="w", width=40)
        self.label_injecting = tk.Label(progress_frame, textvariable=self.injecting_var, anchor="w", width=40)
        self.label_sending = tk.Label(progress_frame, textvariable=self.sending_var, anchor="w", width=40)

        self.label_ping.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.label_progress.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.label_file_list.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.label_processing.grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.label_injecting.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        self.label_sending.grid(row=2, column=1, sticky="w", padx=5, pady=2)

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def load_file_printers(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        file_printers = []
        for printer in printers:
            printer_name = printer[2]
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_details = win32print.GetPrinter(printer_handle, 2)
            win32print.ClosePrinter(printer_handle)
            if printer_details['pPortName'].lower() == "file:":
                file_printers.append(printer_name)

        file_printers.append("Add a New Printer...")
        self.printer_dropdown["values"] = file_printers
        if file_printers:
            self.printer_var.set(file_printers[0])

        self.printer_dropdown.bind("<<ComboboxSelected>>", self.check_add_printer)

    def ping_printer(self):
        ip_address = self.ip_entry.get()

        try:
            ping = subprocess.run(["ping", "-n", "1", ip_address], capture_output=True, text=True, check=True)
            if f"Reply from {ip_address}" in ping.stdout:
                self.ping_var.set("Ping... Success")
            else:
                self.ping_var.set("Ping... Failed")
        except subprocess.CalledProcessError:
            self.ping_var.set("Ping... Failed")

    def check_add_printer(self, event):
        if self.printer_var.get() == "Add a New Printer...":
            os.system("rundll32 printui.dll,PrintUIEntry /il")
            self.load_file_printers()

    def start_process(self):
        directory = self.dir_entry.get()
        printer = self.printer_var.get()
        printer_ip = self.ip_entry.get()

        message = ""

        if not directory:
            message += "Please select directory.\n"
        if not printer:
            message += "Please select FILE printer.\n"
        if not printer_ip:
            message += "Please enter IP Address.\n"
        if "Success" not in self.ping_var:
            message += "Please ping the IP first."

        if message != "":
            messagebox.showwarning("Missing Information", message)
            return

        print_thread = threading.Thread(target=self.thread_process)
        print_thread.start()

    def thread_process(self):
        directory = self.dir_entry.get()
        printer = self.printer_var.get()
        printer_ip = self.ip_entry.get()

        try:
            self.get_file_list()
            self.file_list_size = len(self.file_list)
            self.file_list_var.set("Getting File List... done")

            for file in self.file_list:
                self.file_list_progress += 1
                self.processing_var.set(f"Creating .prn Files... {self.file_list_progress}/{self.file_list_size}")
                db.create_PRN_files(file)
            self.processing_var.set("Creating .prn Files... done")

            self.get_prn_list()
            self.prn_list_size = len(self.prn_list)

            for file in self.prn_list:
                self.prn_list_progress += 1
                self.injecting_var.set(f"Injecting PJL Commands... {self.prn_list_progress}/{self.prn_list_size}")
                db.inject_PJL_commands(file, printer_ip)
                self.sending_var.set(f"Sending File via LPR... {self.prn_list_progress}/{self.prn_list_size}")

            self.injecting_var.set(f"Injecting PJL Commands... done")
            self.sending_var.set(f"Sending File via LPR... done")

        except Exception as e:
            logger.error(e)

        finally:
            pass

    def get_file_list(self):
        directory = Path(self.dir_entry.get())
        excluded_extensions = [".prn"]
        for file_path in directory.rglob("*"):
            if file_path.is_file() and file_path.suffix not in excluded_extensions:
                self.file_list.append(file_path)

    def get_prn_list(self):
        directory = Path(self.dir_entry.get())
        accepted_extensions = [".prn"]
        for file_path in directory.rglob("*"):
            if file_path.is_file() and file_path.suffix in accepted_extensions:
                self.prn_list.append(file_path)


if __name__ == "__main__":
    logger = setup()
    root = tk.Tk()
    app = DocumentBoxGUI(root)
    root.mainloop()
