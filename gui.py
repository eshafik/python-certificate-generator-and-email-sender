import time
import tkinter as tk
from tkinter import simpledialog, IntVar, HORIZONTAL, messagebox
from tkinter.filedialog import askopenfilename
from tkinter.ttk import Progressbar

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

import math

# Popup Window
from sender import send_email


class InputDialog(simpledialog.Dialog):

    def body(self, master):
        self.file_name = None

        self.geometry("400x200")
        tk.Label(master, text="Email Subject:").grid(row=0)
        self.subject = tk.Entry(master)
        self.subject.grid(row=0, column=1)
        tk.Label(master, text="Email Body:").grid(row=1)
        self.body = tk.Entry(master)
        self.body.grid(row=1, column=1)
        tk.Label(master, text="Email Signature:").grid(row=2)
        self.signature = tk.Entry(master)
        self.signature.grid(row=2, column=1)

        tk.Label(master, text="Select Excel File:").grid(row=3)
        self.selected_file = tk.Button(master, text="Browse a file", command=self.file_dialogs)
        self.selected_file.grid(row=3, column=1)

        self.file_path = tk.Label(master, text="")
        self.file_path.grid(row=4, column=1)
        self.check_value = IntVar()
        self.check = tk.Checkbutton(master, text="Send without certificate", variable=self.check_value, onvalue=1,
                                    offvalue=0)
        self.check.grid(row=5, column=1)

        return self.subject

    def apply(self):
        self.subject = self.subject.get()
        self.body = self.body.get()
        self.signature = self.signature.get()
        self.check_value = self.check_value.get()
        if not (self.subject and self.body and self.signature and self.file_name):
            messagebox.showinfo("Error Message", "Please fill up all fields for processing")

    def file_dialogs(self):
        self.filename = askopenfilename(initialdir="/", title="Select a file", filetype=(("excel", ".xlsx"),))
        self.file_name = self.filename
        print("filename: ", self.filename)
        tk.Label(self.file_path, text=self.filename).grid(row=5)


# Popup Window
class EmailProcessingDialog(simpledialog.Dialog):

    def body(self, master):
        self.geometry("400x200")
        self.progress_bar = Progressbar(master, orient=HORIZONTAL, length=100, mode='indeterminate')
        self.progress_bar.grid(row=2, column=1)

        self.button = tk.Button(master, text="Start Sending", command=self.read_excel)
        self.button.grid(row=3, column=1)
        return self.progress_bar

    def read_excel(self):
        try:
            # read excel file
            data = pd.read_excel(result.file_name)

            name_list = data["Name"].tolist()
            id_list = data['ID'].tolist()
            email_list = data['Email'].tolist()
            total_length = len(name_list)

            for name, i, e, idx in zip(name_list, id_list, email_list, range(1, total_length + 1)):
                if not result.check_value:
                    im = Image.open("cert.jpg")
                    d = ImageDraw.Draw(im)
                    location = (320, 216)
                    text_color = (0, 0, 0)
                    font = ImageFont.truetype("CHOPS___.TTF", 50)
                    d.text(location, name, fill=text_color, font=font)
                    location = (405, 284)
                    text_color = (0, 0, 0)
                    font = ImageFont.truetype("CHOPS___.TTF", 25)
                    d.text(location, f"ID: {i}", fill=text_color, font=font)
                    file_name = f"certificate_{i}.pdf"
                    im.save(file_name)
                    send_email(send_to=e, name=name, subject=result.subject, body=result.body, signature=result.signature,
                               attachment=file_name)
                else:
                    send_email(send_to=e, name=name, subject=result.subject, body=result.body, signature=result.signature)
                value = math.ceil(idx * 100 / total_length)
                self.progress_bar['value'] = value
                self.master.update_idletasks()

            tk.Label(self.button, text="Done").grid(row=3)
        except:
            messagebox.showinfo("Error Message", "Something went wrong. Please try again!")


if __name__ == "__main__":
    # Open the window
    root = tk.Tk()
    root.withdraw()
    root.iconphoto(True, tk.PhotoImage(file='icon.png'))
    result = InputDialog(root, "Auto Email Sender")
    if result.file_name:
        processing = EmailProcessingDialog(root, "Auto Email Sending Status")
    # End Popup window
