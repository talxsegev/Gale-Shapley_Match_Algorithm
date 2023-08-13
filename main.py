import tkinter
import tkinter.messagebox
import customtkinter
import time
import pandas as pd
from tkinter import filedialog
import smtplib
from email.mime.text import MIMEText
from validate_email import validate_email
import webbrowser
import urllib.parse
from PIL import Image, ImageTk
from threading import Thread
import cv2





customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        # configure copyright
        # Create the Copyright label
        copyright_label = customtkinter.CTkLabel(self,text="Copyright © Tal Segev, Ofek Rabotnicoff, Barel Cohen, Rachel Kopel, Tal Shemesh")

        # Create the Contact Us link
        contact_us_label = customtkinter.CTkLabel(self, text="Contact Us", text_color=("blue", "blue"), cursor="hand2")
        contact_us_label.bind("<Button-1>", lambda event: self.open_contact_us_email())

        # Place the Copyright label and the Contact Us link in the grid
        copyright_label.grid(row=4, column=0, columnspan=4, pady=(10, 0))
        contact_us_label.grid(row=5, column=0, columnspan=4,pady = (0,10))
        # configure window
        self.title("Matching Application")
        self.geometry(f"{1100}x{580}")

        #icon
        self.iconbitmap('CollegeLogo.ico')

        self.logo_image = tkinter.PhotoImage(file='CollegeLogo.png')
        self.logo_label = tkinter.Label(self, image=self.logo_image, bg="white", compound="right" ,fg="white")
        self.logo_label.grid(row=0, column=2, sticky='ne')  # 'ne' means top-right (north-east)
        # Initial resize
        self.resize_logo()
        # Bind the window resize event
        self.bind('<Configure>', self.resize_logo)

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Menu", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Upload Organizations rating",width=30, command=lambda: self.sidebar_button_event('companies'))
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="Upload student rating",width=30, command=lambda: self.sidebar_button_event('students'))
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        # Set a common size for the buttons using the uniform option
        self.sidebar_frame.grid_columnconfigure(0, uniform="button_column")
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create main entry and button
        self.entry = customtkinter.CTkEntry(self, placeholder_text="Enter your email")
        self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        self.main_button_1 = customtkinter.CTkButton(master=self, text="Send Result by Email", fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), command=self.open_default_email_app)
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")



        # create textbox
        self.textbox = customtkinter.CTkTextbox(self, width=250)
        self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")


        # create slider and progressbar frame
        self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.slider_progressbar_frame.grid(row=1, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
        self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)

        self.calc_button = customtkinter.CTkButton(self.slider_progressbar_frame, text="Calculate Matching", width=30, command=self.start_progressbar)
        self.calc_button.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")

        self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
        self.progressbar_1.grid(row=1, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")

        self.appearance_mode_optionemenu.set("Light")
        self.scaling_optionemenu.set("100%")

        self.textbox.insert("0.0",
                            "Background\n\n" + "The Gale-Shapley algorithm is a stable matching algorithm that ensures each participant is optimally matched based on their preferences, with no incentive to change. It operates by allowing participants to propose and accept or reject proposals, ultimately resulting in stable pairs.\n\n" +
                            "Instructions:\n\n" +
                            "1.Upload Organizations Rating:\n"
                            "Click the \"Upload Organizations Rating\" button and load the CSV file of the scores from the organizations from Google Forms.\n\n" +
                            "2.Upload Student Rating:\n"
                            "Click the \"Upload Student Rating\" button and load the CSV file of the students' scores from Google Forms.\n\n" +
                            "3.Calculate Matching:\n"
                            "Click the \"Calculate Matching\" button and you will see below the matching output of the stable match.\n\n")
        self.result = {}
        self.unmatched_students = []
        self.unmatched_companies = []

        self.geometry("400x400")
        self.title("Main App")

        instruction_button = tk.Button(self, text="Instructions", command=self.open_video_window)
        instruction_button.pack(pady=20)
    def start_progressbar(self):
        self.progressbar_1.configure(mode="indeterminate")
        self.progressbar_1.start()
        self.calculate_matching()  # Start the matching process
        self.after(5000, self.stop_progressbar)  # After 5000 milliseconds (5 seconds), stop the progress bar

    def stop_progressbar(self):
        self.progressbar_1.stop()
        self.progressbar_1.configure(mode="determinate")  # Reset the mode to determinate

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event(self, type):
        if type == 'students':
            self.load_csv('students')
        elif type == 'companies':
            self.load_csv('companies')

    def load_csv(self, type):
        self.textbox.insert("end", f"Attempting to load {type} CSV...\n")
        file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])
        if not file_path:
            self.textbox.insert("end", f"Loading {type} CSV aborted.\n\n")
            self.textbox.see("end")
            return

        try:
            if type == 'students':
                self.students_df = pd.read_csv(file_path)
                self.textbox.insert("end", "Student preferences CSV successfully loaded.\n\n")
            elif type == 'companies':
                self.companies_df = pd.read_csv(file_path)
                self.textbox.insert("end", "Company preferences CSV successfully loaded.\n\n")
        except Exception as e:
            self.textbox.insert("end", f"Error loading {type} CSV: {str(e)}\n")

    def calculate_matching(self):
        self.textbox.delete('1.0', 'end')  # Clear the textbox before inserting results
        self.textbox.insert("end", "Starting matching calculation...\n")

        try:
            student_prefs_dict = {
                row["Full name"]: [col.split('[')[1].split(']')[0] for col in row[1:-1].sort_values().index]
                for _, row in self.students_df.iterrows()}
            self.textbox.insert("end", "Processed student preferences...\n")

            company_prefs_dict = {
                row["Full name"]: [col.split('[')[1].split(']')[0] for col in row[1:-1].sort_values().index]
                for _, row in self.companies_df.iterrows()}
            self.textbox.insert("end", "Processed company preferences...\n")

            self.textbox.insert("end", "Running Gale-Shapley algorithm...\n")
            matchings = self.gale_shapley(student_prefs_dict, company_prefs_dict)
            unmatched_students, unmatched_companies = self.identify_unmatched(matchings, student_prefs_dict,company_prefs_dict)
            self.unmatched_students = unmatched_students
            self.unmatched_companies = unmatched_companies
            self.textbox.insert("end", "Matchings computed. Displaying results...\n")

            # Calculating success rates after matching
            success_rates = self.calculate_success_rate(matchings, student_prefs_dict, company_prefs_dict)
            self.result = success_rates
            self.textbox.insert("end", "\nSuccess Rates:\n")
            for company, data in success_rates.items():
                self.textbox.insert("end",f"{company} matched with {data['student']} has a success rate of {data['success_rate']:.2f}%\n")

            # Display unmatched students and companies
            if unmatched_students:
                self.textbox.insert("end", "\nUnmatched Students:\n")
                for student in unmatched_students:
                    self.textbox.insert("end", f"{student}\n")

            if unmatched_companies:
                self.textbox.insert("end", "\nUnmatched Companies:\n")
                for company in unmatched_companies:
                    self.textbox.insert("end", f"{company}\n")


        except Exception as e:
            self.textbox.insert("end", f"Error: {str(e)}")
        self.textbox.see("end")
    def gale_shapley(self, students_pref, companies_pref):
        """Gale-Shapley algorithm for stable matching."""
        # Initialize
        n = len(students_pref)
        free_students = list(students_pref.keys())
        proposals = {student: [] for student in students_pref}
        matched = {}

        while free_students:
            student = free_students.pop()
            student_prefs = students_pref[student]

            for company in student_prefs:
                # Check if student has already proposed to this company
                if company not in proposals[student]:
                    proposals[student].append(company)

                    # If company is free, accept the proposal
                    if company not in matched:
                        matched[company] = student
                        break

                    # If company is already matched, check if they prefer the new student
                    else:
                        current_match = matched[company]
                        if companies_pref[company].index(student) < companies_pref[company].index(current_match):
                            matched[company] = student
                            free_students.append(current_match)
                            break

        return matched

    def identify_unmatched(self, matched_pairs, students_pref, companies_pref):
        """Identify unmatched students and companies."""
        unmatched_students = [student for student in students_pref if student not in matched_pairs.values()]
        unmatched_companies = [company for company in companies_pref if company not in matched_pairs]

        return unmatched_students, unmatched_companies
    def calculate_success_rate(self, matchings, student_prefs, company_prefs):
        success_rates = {}
        for company, student in matchings.items():
            # Fetch the ranks from the preference lists
            student_rank_for_company = student_prefs[student].index(company) + 1
            company_rank_for_student = company_prefs[company].index(student) + 1
            # Average the ranks
            average_rank = (student_rank_for_company + company_rank_for_student) / 2
            # Derive success rate as the inverse of the average rank
            success_rate =round((1 / average_rank)*100, 2)
            success_rates[company] = {
                "student": student,
                "success_rate": success_rate
            }
        return success_rates

    # def send_email(self, recipient, subject, body):
    #     if not validate_email(recipient):
    #         self.textbox.insert("end", f"'{recipient}' is not a valid email address.\n")
    #         return
    #
    #     # Email setup
    #     sender_email = "your_email@gmail.com"
    #     sender_password = "your_password"
    #
    #     msg = MIMEText(body)
    #     msg['From'] = sender_email
    #     msg['To'] = recipient
    #     msg['Subject'] = subject
    #
    #     # Send the email
    #     try:
    #         with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
    #             server.login(sender_email, sender_password)
    #             server.sendmail(sender_email, recipient, msg.as_string())
    #         self.textbox.insert("end", "Email sent successfully.\n")
    #     except Exception as e:
    #         self.textbox.insert("end", f"Error sending email: {str(e)}\n")
    #
    # def send_results_by_email(self):
    #     recipient = self.entry.get()
    #     subject = "Matching Results"
    #     body = self.textbox.get("1.0", "end-1c")  # Get the contents of the textbox
    #     self.send_email(recipient, subject, body)
    def open_default_email_app(self):
        recipient = self.entry.get()
        subject = "Matching Results"

        # body = self.textbox.get("1.0", "end-1c")  # Get the contents of the textbox

        # Construct the text representation from self.result
        text_content = ""
        for org_name, data in self.result.items():
            text_content += f"Organization: {org_name}%0D%0A"
            text_content += f"Matched Student: {data['student']}%0D%0A"
            text_content += f"Success Rate: {data['success_rate']}%%0D%0A"
            text_content += "-" * 40 + "%0D%0A"  # Separator line

        if self.unmatched_students:
            text_content += str("Unmatched Students:%0D%0A" + str(self.unmatched_students))
        if self.unmatched_companies:
            text_content += str("Unmatched Companies:%0D%0A" + str(self.unmatched_companies))

        if not validate_email(recipient):
            self.textbox.insert("end", f"\n'{recipient}' is not a valid email address.\n")
            self.textbox.see("end")
            return


        body = text_content

        # Construct the mailto URL
        mailto_url = f"mailto:{recipient}?subject={subject}&body={body}"

        # Open the default email application
        webbrowser.open(mailto_url)

    def open_contact_us_email(self):
        recipient = "Matching@gmail.com"
        subject = "Contact Us"
        body = "Please enter your message here."

        # Construct the mailto URL
        mailto_url = f"mailto:{recipient}?subject={subject}&body={body}"

        # Open the default email application
        webbrowser.open(mailto_url)

    def resize_logo(self, event=None):
        # Adjust the logo label's width and height properties based on window size
        new_width = self.winfo_width() // 10  # For example, 10% of window width
        new_height = self.winfo_height() // 10 # For example, 10% of window height

        self.logo_label.configure(width=new_width, height=new_height)

    def open_video_window(self):
        video_win = VideoWindow('path_to_your_video_file.mp4')
        video_win.title("Instruction Video")
        video_win.mainloop()

class VideoWindow(tk.Toplevel):
    def __init__(self, video_path):
        super().__init__()
        self.video_path = video_path
        self.cap = cv2.VideoCapture(self.video_path)
        self.canvas = tk.Canvas(self, width=self.cap.get(cv2.CAP_PROP_FRAME_WIDTH),
                                height=self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        self.canvas.pack()
        self.update_video()

    def update_video(self):
        ret, frame = self.cap.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            photo = ImageTk.PhotoImage(image=Image.fromarray(frame))
            self.canvas.create_image(0, 0, image=photo, anchor=tk.NW)
            self.after(30, self.update_video)

if __name__ == "__main__":
    app = App()
    app.mainloop()