from tkinter import *
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import os
from tkinter.ttk import Combobox
import openpyxl
from openpyxl import Workbook
import pathlib
import tkcalendar

class StudentRegistrationSystem:
    def __init__(self):
        # Set up window
        self.background = "#0C2C55"
        self.framebg = "#EDEDCE"
        self.framefg = "#296374"
        self.color = "#629FAD"

        self.my_wendow = Tk()
        self.my_wendow.title("Student Regristration System")
        self.my_wendow.geometry("1220x650+50+30")
        self.my_wendow.config(bg=self.background)
        self.my_wendow.resizable(False, False)

        # Excel file creation/check
        file = pathlib.Path("student_data.xlsx")
        if not file.exists():
            file = Workbook()
            sheet = file.active
            sheet['A1'] = "Registration No."
            sheet['B1'] = "Name"
            sheet['C1'] = "class"
            sheet['D1'] = "Gender"
            sheet['E1'] = "DOB"
            sheet['F1'] = "Date Of Registration"
            sheet['G1'] = "Religion"
            sheet['H1'] = "Skill"
            sheet['I1'] = "Father Name"
            sheet['J1'] = "Mather Name"
            sheet['K1'] = "Father's Occupation"
            sheet['L1'] = "Mother's Occupation"
            file.save('student_data.xlsx')

        # Ensure Student_Images directory exists
        if not os.path.exists("Student_Images"):
            os.makedirs("Student_Images")

        # Top Frames
        Label(self.my_wendow, text="Email: Parvatcomputertechnology@gmail.com", width=10, height=3, bg="#f0687c", anchor="e").pack(side=TOP, fill=X)
        Label(self.my_wendow, text="STUDENT REGISTATION", width=10, height=2, bg="#C36464", fg='#fff', font='arial 20 bold').pack(side=TOP, fill=X)

        # Search
        self.Search = StringVar()
        Entry(self.my_wendow, textvariable=self.Search, width=15, bd=2, font="arial 20").place(x=830, y=70)
        self.imageicon3 = PhotoImage(file="Images/search.png")
        self.srch = Button(self.my_wendow, text="Search", compound=LEFT, image=self.imageicon3, width=123, bg='#68ddfa', font="arial 13 bold", command=self.search)
        self.srch.place(x=1070, y=66)
        self.imageicon4 = PhotoImage(file="Images/layer.png")
        self.Update_button = Button(self.my_wendow, image=self.imageicon4, bg="#c36464", command=self.Update)
        self.Update_button.place(x=110, y=64)

        # Registration and Date
        Label(self.my_wendow, text="Registration No:", font="arial 13", fg=self.framebg, bg=self.background).place(x=30, y=150)
        Label(self.my_wendow, text="Date:", font="arial 13", fg=self.framebg, bg=self.background).place(x=500, y=150)

        self.Registration = IntVar()
        self.Date = StringVar()

        self.reg_entry = Entry(self.my_wendow, textvariable=self.Registration, width=15, font="arial 10")
        self.reg_entry.place(x=160, y=150)

        self.Date.set(datetime.today().strftime("%d/%m/%Y"))
        self.date_entry = Entry(self.my_wendow, textvariable=self.Date, width=15, font="arial 10")
        self.date_entry.place(x=550, y=150)

        # Student details
        self.obj = LabelFrame(self.my_wendow, text="Student's Datails", font=20, bd=2, width=850, bg=self.framebg, fg=self.framefg, height=205, relief=GROOVE)
        self.obj.place(x=30, y=190)
        Label(self.obj, text="Full Name:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=30, y=30)
        Label(self.obj, text="Date of Birth:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=30, y=80)
        Label(self.obj, text="Gender:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=30, y=130)

        Label(self.obj, text="Class:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=450, y=30)
        Label(self.obj, text="Religion:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=450, y=80)
        Label(self.obj, text="Skills:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=450, y=130)

        self.Name = StringVar()
        self.name_input = Entry(self.obj, textvariable=self.Name, width=20, font="arial 10")
        self.name_input.place(x=160, y=30)

        self.Dob = StringVar()
        self.dob_input = tkcalendar.DateEntry(self.obj, textvariable=self.Dob, width=20, font="arial 10")
        self.dob_input.place(x=160, y=80)

        self.radio = StringVar()
        self.R1 = Radiobutton(self.obj, text="Male", variable=self.radio, value="Male", bg=self.framebg, fg=self.framefg, command=self.selection)
        self.R1.place(x=150, y=130)
        self.R2 = Radiobutton(self.obj, text="Female", variable=self.radio, value="Female", bg=self.framebg, fg=self.framefg, command=self.selection)
        self.R2.place(x=200, y=130)

        self.R1.deselect()
        self.R2.deselect()

        self.R1['indicatoron'] = 1
        self.R2['indicatoron'] = 1

        self.R1.deselect(), self.R2.deselect()

        self.R1_value = None

        self.Religion = StringVar()
        self.religion_input = Entry(self.obj, textvariable=self.Religion, width=20, font="arial 10")
        self.religion_input.place(x=580, y=80)

        self.Skill = StringVar()
        self.skill_input = Entry(self.obj, textvariable=self.Skill, width=20, font="arial 10")
        self.skill_input.place(x=580, y=130)

        self.Class = Combobox(self.obj, values=['ID', 'IDOCS', 'IDOCR', 'PIE', 'PSP', 'MCDPT', 'HB', 'AMOI', 'TMMC',
                                               'HBOMI', 'HBOTHP', 'THI', 'AA', 'AAOCP', 'AAOG', 'TCPS', 'COMPT', 'GE', 'GEOCF',
                                               'GEOCM', 'GEORH', 'DEV', 'DEVOAM', 'DEVOWFS'],
                                               font="Roboto 10", width=17, state="r")
        self.Class.place(x=580, y=30)
        self.Class.set("Select class")

        # Parents details
        self.obj2 = LabelFrame(self.my_wendow, text="Parents's Datails", font=20, bd=2, width=850, bg=self.framebg, fg=self.framefg, height=170, relief=GROOVE)
        self.obj2.place(x=30, y=410)
        Label(self.obj2, text="Father's Name:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=30, y=30)
        Label(self.obj2, text="Occupation:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=30, y=80)
        Label(self.obj2, text="Mother's Name:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=450, y=30)
        Label(self.obj2, text="Occupation:", font="arial 13", bg=self.framebg, fg=self.framefg).place(x=450, y=80)

        self.F_Name = StringVar()
        self.f_input = Entry(self.obj2, textvariable=self.F_Name, width=20, font="arial 10")
        self.f_input.place(x=160, y=30)
        self.Father_Occupation = StringVar()
        self.fo_input = Entry(self.obj2, textvariable=self.Father_Occupation, width=20, font="arial 10")
        self.fo_input.place(x=160, y=80)

        self.M_Name = StringVar()
        self.m_input = Entry(self.obj2, textvariable=self.M_Name, width=20, font="arial 10")
        self.m_input.place(x=580, y=30)
        self.Mother_Occupation = StringVar()
        self.mo_input = Entry(self.obj2, textvariable=self.Mother_Occupation, width=20, font="arial 10")
        self.mo_input.place(x=580, y=80)

        # image
        self.f = Frame(self.my_wendow, bd=3, bg="black", width=200, height=200, relief=GROOVE)
        self.f.place(x=960, y=130)

        self.filename = ""
        try:
            self.image = Image.open("Images/pro.jpg")
            self.img = ImageTk.PhotoImage(self.image)
        except Exception:
            self.image = None
            self.img = None

        self.lbl = Label(self.f, bg="black", image=self.img)
        self.lbl.place(x=0, y=0)

        # buttons
        Button(self.my_wendow, text="Upload", width=19, height=2, font="arial 12 bold", bg="lightblue", command=self.showimage).place(x=960, y=350)
        self.savebutton = Button(self.my_wendow, text="Save", width=19, height=2, font="arial 12 bold", bg="lightgreen", command=self.save)
        self.savebutton.place(x=960, y=420)
        Button(self.my_wendow, text="Reset", width=19, height=2, font="arial 12 bold", bg="lightpink", command=self.clear).place(x=960, y=490)
        Button(self.my_wendow, text="Exit", width=19, height=2, font="arial 12 bold", bg="gray", command=self.Exit).place(x=960, y=560)

        # Load next registration number
        self.registration_no()

        self.my_wendow.mainloop()

    def Exit(self):
        self.my_wendow.destroy()

    def showimage(self):
        try:
            self.filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                        title="Select image file", filetypes=(("JPG File","*.jpg"),
                                                            ("PNG File","*.png"),
                                                            ("All files", "*.txt"),
                                                            ("All files", "*.*")))
            if self.filename:
                self.img_pil = Image.open(self.filename)
                resized_image = self.img_pil.resize((190, 190))
                self.img = ImageTk.PhotoImage(resized_image)
                self.lbl.config(image=self.img)
                self.lbl.image = self.img
                self.image = self.img_pil
            else:
                self.image = None
        except Exception:
            messagebox.showerror("Error", "Unable to open the selected image file.")

    def selection(self):
        value = self.radio.get()
        # Just for compatibility, but self.radio.get() is always a string ("Male"/"Female")
        return value

    def registration_no(self):
        try:
            file = openpyxl.load_workbook('student_data.xlsx')
            sheet = file.active
            row = sheet.max_row
            max_row_value = sheet.cell(row=row, column=1).value
            try:
                self.Registration.set(int(max_row_value) + 1)
            except:
                self.Registration.set("1")
        except Exception:
            self.Registration.set("1")

    def clear(self):
        self.Name.set('')
        self.Dob.set('')
        self.Religion.set('')
        self.Skill.set('')
        self.F_Name.set('')
        self.M_Name.set('')
        self.Father_Occupation.set('')
        self.Mother_Occupation.set('')
        self.Class.set("Select class")
        self.radio.set('')
        self.registration_no()
        self.savebutton.config(state='normal')
        try:
            image1 = Image.open("Images/pro.jpg")
            img1 = ImageTk.PhotoImage(image1)
            self.lbl.config(image=img1)
            self.lbl.image = img1
            self.image = image1
            self.img = img1
        except Exception:
            self.lbl.config(image='')
            self.lbl.image = None
            self.image = None
            self.img = None

    def save(self):
        R1 = self.Registration.get()
        N1 = self.Name.get()
        C1 = self.Class.get()
        G1 = self.radio.get()
        D2 = self.Dob.get()
        D1 = self.Date.get()
        Re1 = self.Religion.get()
        S1 = self.Skill.get()
        fathername = self.F_Name.get()
        mothername = self.M_Name.get()
        F1 = self.Father_Occupation.get()
        M1 = self.Mother_Occupation.get()
        if N1 == "" or C1 == "Select class" or D2 == "" or Re1 == "" or S1 == "" or fathername == "" or mothername == "" or F1 == "" or M1 == "":
            messagebox.showerror("error", "Few Data is missing!")
            return
        if G1 not in ("Male", "Female"):
            messagebox.showerror("error", "Select Gender!")
            return

        try:
            file = openpyxl.load_workbook('student_data.xlsx')
            sheet = file.active
        except Exception:
            messagebox.showerror("File Error", "Unable to open student_data.xlsx for writing.")
            return

        row = sheet.max_row + 1
        sheet.cell(column=1, row=row, value=R1)
        sheet.cell(column=2, row=row, value=N1)
        sheet.cell(column=3, row=row, value=C1)
        sheet.cell(column=4, row=row, value=G1)
        sheet.cell(column=5, row=row, value=D2)
        sheet.cell(column=6, row=row, value=D1)
        sheet.cell(column=7, row=row, value=Re1)
        sheet.cell(column=8, row=row, value=S1)
        sheet.cell(column=9, row=row, value=fathername)
        sheet.cell(column=10, row=row, value=mothername)
        sheet.cell(column=11, row=row, value=F1)
        sheet.cell(column=12, row=row, value=M1)

        try:
            file.save(r'student_data.xlsx')
        except Exception:
            messagebox.showerror("Save Error", "Failed to save Excel file.")
            return

        try:
            if hasattr(self, 'img_pil'):
                self.img_pil.save("Student_Images/" + str(R1) + ".jpg")
            elif hasattr(self, 'image') and self.image:
                self.image.save("Student_Images/" + str(R1) + ".jpg")
        except Exception:
            messagebox.showinfo("info", "Profile Picture is not available!!!!")

        messagebox.showinfo("info", "Sucessfully data Entered!!!!")
        self.clear()
        self.registration_no()

    def search(self):
        text = self.Search.get()
        self.clear()
        self.savebutton.config(state='disabled')
        try:
            num = int(text)
        except Exception:
            messagebox.showerror("Invalid", "Invalid registration number!!!")
            return

        try:
            file = openpyxl.load_workbook('student_data.xlsx')
            sheet = file.active
        except Exception:
            messagebox.showerror("File Error", "Unable to open student_data.xlsx")
            return

        found = False
        row_num = None
        for row in sheet.iter_rows(min_row=2):  # skip headers
            try:
                reg = row[0].value
                if reg is not None and int(reg) == num:
                    found = True
                    row_num = row[0].row
                    break
            except Exception:
                continue

        if not found or not row_num:
            messagebox.showerror("Invalid", "Invalid registration number!!!")
            return

        x1 = sheet.cell(row=row_num, column=1).value
        x2 = sheet.cell(row=row_num, column=2).value
        x3 = sheet.cell(row=row_num, column=3).value
        x4 = sheet.cell(row=row_num, column=4).value
        x5 = sheet.cell(row=row_num, column=5).value
        x6 = sheet.cell(row=row_num, column=6).value
        x7 = sheet.cell(row=row_num, column=7).value
        x8 = sheet.cell(row=row_num, column=8).value
        x9 = sheet.cell(row=row_num, column=9).value
        x10 = sheet.cell(row=row_num, column=10).value
        x11 = sheet.cell(row=row_num, column=11).value
        x12 = sheet.cell(row=row_num, column=12).value

        self.Registration.set(x1)
        self.Name.set(x2)
        self.Class.set(x3)
        self.radio.set(x4)
        if x4 == "Female":
            self.R2.select()
        elif x4 == "Male":
            self.R1.select()
        else:
            self.R1.deselect(), self.R2.deselect()
        self.Dob.set(x5)
        self.Date.set(x6)
        self.Religion.set(x7)
        self.Skill.set(x8)
        self.F_Name.set(x9)
        self.M_Name.set(x10)
        self.Father_Occupation.set(x11)
        self.Mother_Occupation.set(x12)
        img_path = "Student_Images/" + str(x1) + ".jpg"
        try:
            if os.path.exists(img_path):
                img = Image.open(img_path)
                resized_image = img.resize((190, 190))
                photo2 = ImageTk.PhotoImage(resized_image)
                self.lbl.config(image=photo2)
                self.lbl.image = photo2
                self.image = img
                self.img = photo2
            else:
                # use default
                image1 = Image.open("Images/pro.jpg")
                img1 = ImageTk.PhotoImage(image1)
                self.lbl.config(image=img1)
                self.lbl.image = img1
        except Exception:
            pass

    def Update(self):
        R1 = self.Registration.get()
        N1 = self.Name.get()
        C1 = self.Class.get()
        G1 = self.radio.get()
        D2 = self.Dob.get()
        D1 = self.Date.get()
        Re1 = self.Religion.get()
        S1 = self.Skill.get()
        fathername = self.F_Name.get()
        mothername = self.M_Name.get()
        F1 = self.Father_Occupation.get()
        M1 = self.Mother_Occupation.get()
        try:
            file = openpyxl.load_workbook('student_data.xlsx')
            sheet = file.active
        except Exception:
            messagebox.showerror("File Error", "Could not open Excel file for update.")
            return
        found = False
        row_num = None
        for row in sheet.iter_rows(min_row=2):
            try:
                reg = row[0].value
                if reg is not None and int(reg) == int(R1):
                    found = True
                    row_num = row[0].row
                    break
            except Exception:
                continue
        if not found or not row_num:
            messagebox.showerror("Update", "Registration number not found.")
            return

        sheet.cell(column=1, row=int(row_num), value=R1)
        sheet.cell(column=2, row=int(row_num), value=N1)
        sheet.cell(column=3, row=int(row_num), value=C1)
        sheet.cell(column=4, row=int(row_num), value=G1)
        sheet.cell(column=5, row=int(row_num), value=D2)
        sheet.cell(column=6, row=int(row_num), value=D1)
        sheet.cell(column=7, row=int(row_num), value=Re1)
        sheet.cell(column=8, row=int(row_num), value=S1)
        sheet.cell(column=9, row=int(row_num), value=fathername)
        sheet.cell(column=10, row=int(row_num), value=mothername)
        sheet.cell(column=11, row=int(row_num), value=F1)
        sheet.cell(column=12, row=int(row_num), value=M1)
        file.save('student_data.xlsx')
        try:
            if hasattr(self, 'img_pil'):
                self.img_pil.save("Student_Images/" + str(R1) + ".jpg")
            elif hasattr(self, 'image') and self.image:
                self.image.save("Student_Images/" + str(R1) + ".jpg")
        except Exception:
            pass
        messagebox.showinfo("Update", "Update Sucessfully!!")
        self.clear()

if __name__ == "__main__":
    app = StudentRegistrationSystem()
