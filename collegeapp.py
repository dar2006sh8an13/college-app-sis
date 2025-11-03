import tkinter as tk
from tkinter import ttk, messagebox, Menu, scrolledtext
import json
import os
import hashlib
import uuid
from datetime import datetime

# --- NEW: Import openpyxl for Excel export ---
try:
    import openpyxl
except ImportError:
    messagebox.showerror("Missing Library", 
                         "The 'openpyxl' library is required for Excel export.\n\nPlease install it by running:\n\n'pip install openpyxl'")
    exit()


# --- Configuration and Data Files ---
USERS_FILE = 'users.json'
STUDENT_DATA_FILE = 'student_data.json'
SESSION_FILE = 'session.json' # For Remember Me
APP_CODE = "SVCE2024AI" # App code for Admin creation

# Hardcoded student list for Teacher Dashboard view
STUDENT_ROLL = {
    "1sv24ai001": "AAFIYA", "1sv24ai002": "ABHILASH", "1sv24ai003": "AMULYA_E",
    "1sv24ai004": "ANURAG_S_BEDRE", "1sv24ai005": "ARBIYA_BANU", "1sv24ai006": "BHOOMIKA_T",
    "1sv24ai007": "CHANDRASEKHAR_K", "1sv24ai008": "CHARANKUMAR_C_GANIGER", "1sv24ai009": "D_MANJULA",
    "1sv24ai010": "DARSHAN_PL", "1sv24ai011": "DHANALAKSHMI_R", "1sv24ai012": "DIVYA_N",
    "1sv24ai013": "DIVYASHREE_N_V", "1sv24ai014": "EILEEN_PRASAD", "1sv24ai015": "FIZA",
    "1sv24ai016": "GANESH_R_KULKARNI", "1sv24ai017": "GURUNATH_PRASHANT_BANKAPUR",
    "1sv24ai018": "H_M_GANESH", "1sv24ai019": "J_BHUVANESH", "1sv24ai020": "JEEVAN_REDDY",
    "1sv24ai021": "KAIVALYA_REDDY_Y_P", "1sv22ai022": "KANTESH_BIJAPUR", "1sv24ai023": "LAXMI_KURUVATTIGOUDA",
    "1sv24ai024": "LIKHITHA_N", "1sv24ai025": "M_THRUPTHI", "1sv24ai026": "MALLIKARJUN_IRAPPA_CHITTI",
    "1sv24ai027": "MALLIKARJUN_SIDDAPPA", "1sv24ai028": "NIKHIL_H_O", "1sv24ai029": "NITIN_S",
    "1sv24ai030": "PRAJWAL_S_S", "1sv24ai031": "RAHUL_G", "1sv24ai032": "RAHUL_P",
    "1sv24ai033": "RAKSHITHA_D_J", "1sv24ai034": "RAZIYA_SULTANA", "1sv24ai035": "RUMAN_K",
    "1sv24ai036": "S_J_GANESHGOUDA", "1sv24ai037": "SADASHIV_SANTOSHKUMAR", "1sv24ai038": "SHAIK_MOHAMMED_NAQEEB",
    "1sv24ai039": "SHARANYA_N_R", "1sv24ai040": "SHASHANK_M_L", "1sv24ai041": "SIDDALINGAYYA_R_ARADANAMAT",
    "1sv24ai042": "SIDDESHWAR_SURESH_HEGALE", "1sv24ai043": "SRUJAN_R_P", "1sv24ai044": "SUFIYA_KOUSER",
    "1sv24ai045": "SUKANYA_SHRISHAILA", "1sv24ai047": "VIKAS_AWATI", "1sv24ai048": "VINAYAKA_V",
    "dhanush r": "DHANUSH_R", "m vijaykumar": "M_VIJAYKUMAR", "priya k m": "PRIYA_K_M",
    "ranjitha p": "RANJITHA_P", "tharun kumar g": "THARUN_KUMAR_G"
}

# --- Modern Message Popup ---

class MessagePopup(tk.Toplevel):
    """A 'modern' modal popup to replace standard messagebox."""
    def __init__(self, parent, title="Alert", message="", msg_type="info"):
        super().__init__(parent)
        self.title(title)
        
        # Calculate position to center on parent
        parent.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        width = 300
        height = 120
        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.resizable(False, False)
        
        # Set background color for the whole window
        self.configure(bg="#F8F8F8")

        # Color bar
        color = "#00796B" if msg_type == "info" else "#D32F2F"
        color_bar = tk.Frame(self, bg=color, height=10)
        color_bar.pack(fill='x', side=tk.TOP)
        
        # Message
        msg_frame = ttk.Frame(self, padding="15")
        msg_frame.pack(fill='both', expand=True)
        
        ttk.Label(msg_frame, text=message, wraplength=270, justify=tk.LEFT, anchor='w').pack(fill='x', expand=True)
        
        # OK Button
        btn_frame = ttk.Frame(self, padding=(0, 0, 15, 10))
        btn_frame.pack(fill='x', side=tk.BOTTOM)
        
        ok_btn = ttk.Button(btn_frame, text="OK", command=self.destroy, style='TButton')
        ok_btn.pack(side=tk.RIGHT)
        
        # Make modal
        self.transient(parent)
        self.grab_set()
        self.wait_window()

# --- Utility Functions for Data Handling ---

def hash_password(password):
    """Hashes the password for secure storage."""
    return hashlib.sha256(password.encode()).hexdigest()

def load_data(filename):
    """Loads data from a JSON file. Creates an empty file if it doesn't exist."""
    if not os.path.exists(filename):
        with open(filename, 'w') as f:
            if filename == USERS_FILE:
                json.dump({}, f, indent=4)
                return {}
            elif filename == SESSION_FILE:
                json.dump({}, f, indent=4)
                return {}
            else:
                json.dump([], f, indent=4)
                return []
    
    with open(filename, 'r') as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            # Handle case where file might be empty or corrupted
            if filename == USERS_FILE or filename == SESSION_FILE:
                return {}
            else:
                return []

def save_data(filename, data, controller):
    """Saves data to a JSON file."""
    try:
        with open(filename, 'w') as f:
            json.dump(data, f, indent=4)
    except (IOError, PermissionError) as e:
        # Use the controller's root window as the parent for the popup
        MessagePopup(controller.root, 
                     title="Save Error", 
                     message=f"Fatal Error: Could not save data to {filename}.\nReason: {e}\nPlease check file permissions.", 
                     msg_type="error")

# --- Mark Calculation Logic ---
def calculate_marks(record):
    """Calculates internal and total marks based on a student record."""
    i_marks = [record.get(f'i{i}', 0) for i in range(1, 4)]
    a_marks = [record.get(f'a{i}', 0) for i in range(1, 4)]
    lab_marks = record.get('lab_marks', 0)
    external_mark = record.get('external_mark', 0)
    
    best_two_internals = sum(i_marks) - min(i_marks)
    scaled_internal = round((best_two_internals / 40) * 20) if best_two_internals else 0
    
    total_assignment = sum(a_marks)
    scaled_assignment = round((total_assignment / 15) * 5) if total_assignment else 0
    
    scaled_lab = lab_marks
    
    total_internal_mark = scaled_internal + scaled_assignment + scaled_lab
    total_mark = total_internal_mark + external_mark
    
    return total_internal_mark, external_mark, total_mark

# --- Main Application Class ---

class SISApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Student Information System (SIS)")
        self.geometry("900x700") # Increased window size for new layouts
        self.configure(bg="#F8F8F8")
        
        # Provide a reference to the root window for popups
        self.root = self
        
        # Load data
        self.users = load_data(USERS_FILE)
        self.student_data = load_data(STUDENT_DATA_FILE)
        
        self.current_user = None
        
        # Setup the main container frame
        container = ttk.Frame(self, padding="20")
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.frames = {}
        
        # Initialize all pages
        for F in (RoleSelectionPage, StudentLoginPage, GenericLoginPage, ForgotPasswordPage,
                  StudentSignupPage, TeacherSignupPage, AdminSignupPage, 
                  StudentDashboard, TeacherDashboard, TeacherMarksEntryPage, 
                  TeacherAttendanceEntryPage, AdminDashboard, AdminCreationPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Set theme and style
        style = ttk.Style(self)
        style.theme_use('clam') # Use a simple, modern theme
        style.configure('TFrame', background='#F8F8F8')
        style.configure('TLabel', background='#F8F8F8', font=('Arial', 10))
        style.configure('Header.TLabel', font=('Arial', 16, 'bold'), foreground='#004D40') # Dark Green
        style.configure('Big.TButton', font=('Arial', 12, 'bold'), padding=(10, 15))
        style.configure('TButton', font=('Arial', 10, 'bold'), background='#00796B', foreground='white')
        style.map('TButton', background=[('active', '#009688')])
        style.configure('Link.TButton', font=('Arial', 9, 'underline'), 
                        background='#F8F8F8', foreground='#00695C', borderwidth=0, relief='flat')
        style.map('Link.TButton', 
                  foreground=[('active', '#004D40')], 
                  background=[('active', '#F8F8F8')])
        
        self.try_auto_login() # Check for saved session

    def show_frame(self, page_name, **kwargs):
        """Switches the view to the desired frame."""
        frame = self.frames[page_name]
        
        # Pass data to the frame if needed (e.g., setting the role for GenericLoginPage)
        if hasattr(frame, 'on_show'):
            frame.on_show(**kwargs)
            
        frame.tkraise()
        
        # Call refresh method if available
        if hasattr(frame, 'refresh_data'):
            frame.refresh_data()

    # --- Session/Login Functions ---
    
    def save_session(self, role, identifier):
        """Saves user session for 'Remember Me'."""
        session_data = {'role': role, 'identifier': identifier}
        save_data(SESSION_FILE, session_data, self)
        
    def load_session(self):
        """Loads user session."""
        return load_data(SESSION_FILE)
        
    def clear_session(self):
        """Clears the saved session."""
        save_data(SESSION_FILE, {}, self)

    def try_auto_login(self):
        """Checks for a saved session and logs in automatically."""
        session = self.load_session()
        if session and 'role' in session and 'identifier' in session:
            role = session['role']
            identifier = session['identifier']
            
            user_data = None
            if role == 'student':
                user_data = self.get_user_by_usn(identifier)
            else:
                user_data = self.users.get(identifier) # Admin/Teacher use username as identifier
            
            if user_data and user_data['role'] == role:
                self.current_user = user_data
                MessagePopup(self, "Login Success", f"Welcome back, {self.current_user['name']}!", msg_type="info")
                
                if role == 'student':
                    self.show_frame("StudentDashboard")
                elif role == 'teacher':
                    self.show_frame("TeacherDashboard")
                elif role == 'admin':
                    self.setup_admin_menu()
                    self.show_frame("AdminDashboard")
                return

        # If auto-login fails or no session, show role selection
        self.show_frame("RoleSelectionPage")

    def login(self, username, password, code, role_to_check, remember_me):
        """Authenticates a Teacher or Admin based on username, password, and code."""
        hashed_password = hash_password(password)
        user_data = self.users.get(username)

        if not user_data:
            return False

        # Check common fields
        if user_data['role'] != role_to_check or user_data['password'] != hashed_password:
            return False

        # Check role-specific code
        if role_to_check == 'teacher' and user_data.get('subject_code', '').lower() != code.lower():
            return False
        
        if role_to_check == 'admin' and user_data.get('dept_code', '').lower() != code.lower():
            return False

        # All checks passed
        self.current_user = user_data
        
        if remember_me:
            self.save_session(role_to_check, username) # Use username as identifier
        else:
            self.clear_session()
            
        MessagePopup(self, "Login Success", f"Welcome, {self.current_user['name']}!", msg_type="info")
        
        if role_to_check == 'teacher':
            self.show_frame("TeacherDashboard")
        elif role_to_check == 'admin':
            self.setup_admin_menu()
            self.show_frame("AdminDashboard")
            
        return True

    def login_by_usn(self, usn, password, remember_me):
        """Authenticates a student based on USN."""
        hashed_password = hash_password(password)
        
        user_data = self.get_user_by_usn(usn)
        
        if user_data and user_data['password'] == hashed_password:
            self.current_user = user_data
            
            if remember_me:
                self.save_session('student', usn) # Use USN as identifier
            else:
                self.clear_session()
                
            MessagePopup(self, "Login Success", f"Welcome, {self.current_user['name']}!", msg_type="info")
            self.show_frame("StudentDashboard")
            return True
        return False

    def logout(self):
        """Logs out the current user and returns to the login page."""
        self.current_user = None
        self.clear_session() # Clear 'Remember Me'
        
        # Remove Admin menu if it exists
        if hasattr(self, 'admin_menu_bar'):
            self.config(menu=None)
            del self.admin_menu_bar
            
        MessagePopup(self, "Logout", "You have been logged out.", msg_type="info")
        self.show_frame("RoleSelectionPage")

    def get_user_by_usn(self, usn):
        """Finds a student user by USN."""
        for user in self.users.values():
            if user['role'] == 'student' and user.get('usn', '').lower() == usn.lower():
                return user
        return None

    def reset_password(self, identifier, email, role):
        """Resets a user's password to 'password123' if identifier and email match."""
        
        user_to_reset = None
        username_to_reset = None
        
        if role == 'student':
            user_to_reset = self.get_user_by_usn(identifier)
            if user_to_reset:
                username_to_reset = user_to_reset['username']
        else: # admin or teacher
            user_to_reset = self.users.get(identifier)
            if user_to_reset:
                username_to_reset = identifier

        if user_to_reset and user_to_reset.get('email', '').lower() == email.lower():
            new_pass = "password123"
            new_hashed_pass = hash_password(new_pass)
            
            # Update the user in self.users
            self.users[username_to_reset]['password'] = new_hashed_pass
            save_data(USERS_FILE, self.users, self)
            
            MessagePopup(self, "Password Reset", 
                         "Your password has been reset to 'password123'. Please log in.", 
                         msg_type="info")
            self.show_frame("RoleSelectionPage")
        else:
            MessagePopup(self, "Error", 
                         "No account found matching that information. Please check your details.", 
                         msg_type="error")


    def setup_admin_menu(self):
        """Creates the 'Iceberg Menu' for the Admin Dashboard."""
        self.admin_menu_bar = Menu(self)
        self.config(menu=self.admin_menu_bar)
        
        admin_menu = Menu(self.admin_menu_bar, tearoff=0)
        self.admin_menu_bar.add_cascade(label="Admin Actions", menu=admin_menu)
        admin_menu.add_command(label="Create New Admin", command=lambda: self.show_frame("AdminCreationPage"))
        admin_menu.add_separator()
        admin_menu.add_command(label="Logout", command=self.logout)

    # --- NEW: Excel Export Function ---
    def export_to_excel(self):
        """Exports all student data (marks and attendance) to an Excel file."""
        if not self.current_user or self.current_user['role'] not in ('teacher', 'admin'):
             MessagePopup(self, "Error", "You do not have permission to export.", msg_type="error")
             return

        try:
            wb = openpyxl.Workbook()
            
            # --- Marks Sheet ---
            ws_marks = wb.active
            ws_marks.title = "Marks"
            
            # Get subject code for teacher
            teacher_subject_code = self.current_user.get('subject_code', 'ALL')
            
            mark_headers = ["USN", "Name", "Subject", "I1", "I2", "I3", "A1", "A2", "A3", "Lab", "External", "Total Internal", "Total Mark"]
            ws_marks.append(mark_headers)
            
            # --- Attendance Sheet ---
            ws_attendance = wb.create_sheet(title="Attendance")
            attendance_headers = ["USN", "Name", "Subject", "Total Classes", "Attended Classes", "Attendance %"]
            ws_attendance.append(attendance_headers)

            # Prepare user lookups
            student_names = {user['usn'].upper(): user['name'] 
                             for user in self.users.values() if user['role'] == 'student'}

            # Populate data
            for record in self.student_data:
                usn = record.get('usn')
                subject_code = record.get('subject_code')
                
                # If teacher, only export their subject
                if self.current_user['role'] == 'teacher' and subject_code != teacher_subject_code:
                    continue
                    
                name = student_names.get(usn, 'N/A')
                
                # --- Add to Marks Sheet ---
                total_internal, external, total_mark = calculate_marks(record)
                ws_marks.append([
                    usn, name, subject_code,
                    record.get('i1', 0), record.get('i2', 0), record.get('i3', 0),
                    record.get('a1', 0), record.get('a2', 0), record.get('a3', 0),
                    record.get('lab_marks', 0), external,
                    total_internal, total_mark
                ])
                
                # --- Add to Attendance Sheet ---
                total_classes = record.get('total_classes', 0)
                attended_classes = record.get('attended_classes', 0)
                percentage = (attended_classes / total_classes * 100) if total_classes > 0 else 0
                
                ws_attendance.append([
                    usn, name, subject_code,
                    total_classes, attended_classes, f"{percentage:.2f}%"
                ])

            # Save the file
            filename = f"student_data_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb.save(filename)
            MessagePopup(self, "Success", f"Data exported successfully to:\n{filename}", msg_type="info")
            
        except Exception as e:
            MessagePopup(self, "Export Error", f"An error occurred during export:\n{e}", msg_type="error")


# --- Base Page Class ---

class BasePage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

# --- Login Pages (Refactored) ---

class RoleSelectionPage(BasePage):
    """New starting page to select user role before logging in."""
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="40")
        
        center_frame = ttk.Frame(self)
        center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        ttk.Label(center_frame, text="Student Information System", style='Header.TLabel').grid(row=0, column=0, columnspan=2, pady=20)
        
        ttk.Label(center_frame, text="Please select your role to log in:", font=('Arial', 11)).grid(row=1, column=0, columnspan=2, pady=(10, 20))
        
        ttk.Button(center_frame, text="Student Login", style='Big.TButton',
                   command=lambda: controller.show_frame("StudentLoginPage")).grid(row=2, column=0, columnspan=2, pady=5, sticky="ew")
        
        ttk.Button(center_frame, text="Teacher Login", style='Big.TButton',
                   command=lambda: controller.show_frame("GenericLoginPage", role='teacher')).grid(row=3, column=0, columnspan=2, pady=5, sticky="ew")
        
        ttk.Button(center_frame, text="Admin Login", style='Big.TButton',
                   command=lambda: controller.show_frame("GenericLoginPage", role='admin')).grid(row=4, column=0, columnspan=2, pady=5, sticky="ew")

        # Separator
        ttk.Separator(center_frame, orient='horizontal').grid(row=5, column=0, columnspan=2, sticky="ew", pady=20)

        # Signup Section
        ttk.Label(center_frame, text="Don't have an account?").grid(row=6, column=0, columnspan=2, pady=(10, 5))
        
        signup_frame = ttk.Frame(center_frame)
        signup_frame.grid(row=7, column=0, columnspan=2, pady=5)
        
        ttk.Button(signup_frame, text="Student Signup", 
                   command=lambda: controller.show_frame("StudentSignupPage")).pack(side=tk.LEFT, padx=5)
        ttk.Button(signup_frame, text="Teacher Signup", 
                   command=lambda: controller.show_frame("TeacherSignupPage")).pack(side=tk.LEFT, padx=5)
        ttk.Button(signup_frame, text="Admin Signup (Initial)", 
                   command=lambda: controller.show_frame("AdminSignupPage")).pack(side=tk.LEFT, padx=5)

class StudentLoginPage(BasePage):
    """Dedicated login page for students using USN."""
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="40")
        
        self.remember_me_var = tk.BooleanVar()
        
        center_frame = ttk.Frame(self)
        center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        self.title_label = ttk.Label(center_frame, text="Student Login", style='Header.TLabel')
        self.title_label.grid(row=0, column=0, columnspan=2, pady=20)
        
        # USN Label
        ttk.Label(center_frame, text="USN:", anchor='w').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.usn_entry = ttk.Entry(center_frame, width=30)
        self.usn_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Password Label
        ttk.Label(center_frame, text="Password:", anchor='w').grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.password_entry = ttk.Entry(center_frame, show="*", width=30)
        self.password_entry.grid(row=2, column=1, padx=5, pady=5)

        # Remember Me & Forgot Password
        options_frame = ttk.Frame(center_frame)
        options_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky='ew')
        
        ttk.Checkbutton(options_frame, text="Remember Me", variable=self.remember_me_var).pack(side=tk.LEFT)
        
        ttk.Button(options_frame, text="Forgot Password?", style='Link.TButton',
                   command=lambda: controller.show_frame("ForgotPasswordPage", role='student')).pack(side=tk.RIGHT)

        # Login Button
        login_button = ttk.Button(center_frame, text="Login", 
                                  command=self._login_command)
        login_button.grid(row=4, column=0, columnspan=2, pady=(10, 10), sticky="ew")

        # Back Button
        back_button = ttk.Button(center_frame, text="Back to Role Selection", 
                                 command=lambda: controller.show_frame("RoleSelectionPage"))
        back_button.grid(row=5, column=0, columnspan=2, pady=10)

    def _login_command(self):
        usn = self.usn_entry.get().strip()
        password = self.password_entry.get().strip()
        remember_me = self.remember_me_var.get()
        
        if not usn or not password:
            MessagePopup(self, "Error", "Please fill in all fields.", msg_type="error")
            return

        if self.controller.login_by_usn(usn, password, remember_me):
            self.usn_entry.delete(0, tk.END)
            self.password_entry.delete(0, tk.END)
        else:
            MessagePopup(self, "Login Failed", "Invalid USN or Password.", msg_type="error")

class GenericLoginPage(BasePage):
    """Login page for Teacher and Admin using Username, Code, and Password."""
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="40")
        self.role = 'teacher' # Default
        self.remember_me_var = tk.BooleanVar()
        
        center_frame = ttk.Frame(self)
        center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        self.title_label = ttk.Label(center_frame, text="Login", style='Header.TLabel')
        self.title_label.grid(row=0, column=0, columnspan=2, pady=20)
        
        # Username Label
        ttk.Label(center_frame, text="Username:", anchor='w').grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.username_entry = ttk.Entry(center_frame, width=30)
        self.username_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Code Label (Subject or Dept)
        self.code_label = ttk.Label(center_frame, text="Code:", anchor='w')
        self.code_label.grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.code_entry = ttk.Entry(center_frame, width=30)
        self.code_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # Password Label
        ttk.Label(center_frame, text="Password:", anchor='w').grid(row=3, column=0, sticky='w', padx=5, pady=5)
        self.password_entry = ttk.Entry(center_frame, show="*", width=30)
        self.password_entry.grid(row=3, column=1, padx=5, pady=5)
        
        # Remember Me & Forgot Password
        options_frame = ttk.Frame(center_frame)
        options_frame.grid(row=4, column=0, columnspan=2, pady=5, sticky='ew')
        
        ttk.Checkbutton(options_frame, text="Remember Me", variable=self.remember_me_var).pack(side=tk.LEFT)
        
        self.forgot_btn = ttk.Button(options_frame, text="Forgot Password?", style='Link.TButton',
                   command=lambda: controller.show_frame("ForgotPasswordPage", role=self.role))
        self.forgot_btn.pack(side=tk.RIGHT)

        # Login Button
        login_button = ttk.Button(center_frame, text="Login", 
                                  command=self._login_command)
        login_button.grid(row=5, column=0, columnspan=2, pady=(10, 10), sticky="ew")
        
        # Back Button
        back_button = ttk.Button(center_frame, text="Back to Role Selection", 
                                 command=lambda: controller.show_frame("RoleSelectionPage"))
        back_button.grid(row=6, column=0, columnspan=2, pady=10)

    def on_show(self, role):
        """Called when the frame is shown to set the role."""
        self.role = role
        if role == 'teacher':
            self.title_label.config(text="Teacher Login")
            self.code_label.config(text="Subject Code:")
        elif role == 'admin':
            self.title_label.config(text="Admin Login")
            self.code_label.config(text="Department Code:")
        
        # Update forgot password button command
        self.forgot_btn.config(command=lambda: self.controller.show_frame("ForgotPasswordPage", role=self.role))
            
        # Clear fields on show
        self.username_entry.delete(0, tk.END)
        self.code_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)
        self.remember_me_var.set(False)


    def _login_command(self):
        """Handles the login process."""
        username = self.username_entry.get().strip()
        code = self.code_entry.get().strip()
        password = self.password_entry.get().strip()
        remember_me = self.remember_me_var.get()
        
        if not username or not code or not password:
            MessagePopup(self, "Error", "Please fill in all fields.", msg_type="error")
            return

        if self.controller.login(username, password, code, self.role, remember_me):
            self.username_entry.delete(0, tk.END)
            self.code_entry.delete(0, tk.END)
            self.password_entry.delete(0, tk.END)
        else:
            MessagePopup(self, "Login Failed", "Invalid credentials. Please check all fields.", msg_type="error")

# --- Forgot Password Page ---

class ForgotPasswordPage(BasePage):
    """Page to handle password reset logic."""
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="40")
        self.role = 'student' # default
        
        center_frame = ttk.Frame(self)
        center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        self.title_label = ttk.Label(center_frame, text="Reset Password", style='Header.TLabel')
        self.title_label.grid(row=0, column=0, columnspan=2, pady=20)
        
        # Identifier Label (USN or Username)
        self.id_label = ttk.Label(center_frame, text="Identifier:", anchor='w')
        self.id_label.grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.id_entry = ttk.Entry(center_frame, width=30)
        self.id_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Email Label
        ttk.Label(center_frame, text="Registered Email:", anchor='w').grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.email_entry = ttk.Entry(center_frame, width=30)
        self.email_entry.grid(row=2, column=1, padx=5, pady=5)

        # Reset Button
        reset_button = ttk.Button(center_frame, text="Reset Password", 
                                  command=self._reset_command)
        reset_button.grid(row=3, column=0, columnspan=2, pady=20, sticky="ew")

        # Back Button
        back_button = ttk.Button(center_frame, text="Back to Role Selection", 
                                 command=lambda: controller.show_frame("RoleSelectionPage"))
        back_button.grid(row=4, column=0, columnspan=2, pady=10)

    def on_show(self, role):
        """Set the role and update labels."""
        self.role = role
        if role == 'student':
            self.id_label.config(text="USN:")
        else:
            self.id_label.config(text="Username:")
        
        # Clear fields
        self.id_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)

    def _reset_command(self):
        identifier = self.id_entry.get().strip()
        email = self.email_entry.get().strip()
        
        if not identifier or not email:
            MessagePopup(self, "Error", "Please fill in all fields.", msg_type="error")
            return

        self.controller.reset_password(identifier, email, self.role)


# --- Signup Pages ---

class StudentSignupPage(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="20")
        
        ttk.Label(self, text="Student Signup", style='Header.TLabel').pack(pady=10)
        
        form_frame = ttk.Frame(self)
        form_frame.pack(pady=10)
        
        fields = [
            ("Name", "name", "entry"), 
            ("USN", "usn", "entry"),
            ("Sem", "sem", "entry"),
            ("Gender", "gender", "combo"),
            ("Phone Number", "phone", "entry"), 
            ("Email ID", "email", "entry"), 
            ("Department Name", "dept_name", "entry"), # Added Dept Name
            ("Password", "password", "entry_pass")
        ]
        
        self.entries = {}
        for i, (label_text, key, widget_type) in enumerate(fields):
            ttk.Label(form_frame, text=label_text + ":").grid(row=i, column=0, sticky='w', padx=5, pady=5)
            
            if widget_type == "entry":
                entry = ttk.Entry(form_frame, width=40)
            elif widget_type == "entry_pass":
                entry = ttk.Entry(form_frame, width=40, show="*")
            elif widget_type == "combo":
                entry = ttk.Combobox(form_frame, width=38, values=["Male", "Female", "Other"], state="readonly")
            
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.entries[key] = entry
            
        ttk.Button(self, text="Register Student", command=self._register_command).pack(pady=15)
        ttk.Button(self, text="Back to Login", command=lambda: controller.show_frame("RoleSelectionPage")).pack()

    def _register_command(self):
        """Registers a new student user."""
        data = {key: entry.get().strip() for key, entry in self.entries.items()}
        
        required_fields = ["name", "usn", "sem", "gender", "phone", "email", "dept_name", "password"]
        if any(not data[f] for f in required_fields):
            MessagePopup(self, "Error", "All fields are required.", msg_type="error")
            return

        if data['usn'].lower() not in STUDENT_ROLL:
            MessagePopup(self, "Error", f"Invalid USN: {data['usn']}\n\nThis USN is not on the official department roll.", msg_type="error")
            return

        student_username = data['usn'].lower()
        
        if student_username in self.controller.users:
            MessagePopup(self, "Error", "Username (USN) already taken.", msg_type="error")
            return

        new_user = {
            'role': 'student',
            'name': data['name'],
            'usn': data['usn'].upper(),
            'sem': data['sem'],
            'gender': data['gender'],
            'phone': data['phone'],
            'email': data['email'],
            'dept_name': data['dept_name'],
            'username': student_username,
            'password': hash_password(data['password']),
            'created_at': datetime.now().isoformat()
        }
        
        self.controller.users[student_username] = new_user
        save_data(USERS_FILE, self.controller.users, self.controller)
        
        MessagePopup(self, "Success", "Student registered successfully. You can now log in.", msg_type="info")
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.controller.show_frame("RoleSelectionPage")

class TeacherSignupPage(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="20")
        
        ttk.Label(self, text="Teacher Signup", style='Header.TLabel').pack(pady=10)
        
        self.form_frame = ttk.Frame(self)
        self.form_frame.pack(pady=10)
        
        self.fields_frame = ttk.Frame(self.form_frame)
        self.fields_frame.pack()

        fields = [
            ("Name", "name_entry"), 
            ("Subject Name", "subject_name_entry"), 
            ("Subject Code", "subject_code_entry"), 
            ("Phone Number", "phone_entry"), 
            ("Email ID", "email_entry"), 
            ("Department Name", "dept_name_entry"), # Added Dept Name
            ("Username", "username_entry"),
            ("Password", "password_entry", True)
        ]
        
        self.entries = {}
        for i, (label_text, key, *is_password) in enumerate(fields):
            label = ttk.Label(self.fields_frame, text=label_text + ":")
            label.grid(row=i, column=0, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(self.fields_frame, width=40, show="*" if is_password else "")
            entry.grid(row=i, column=1, padx=5, pady=5)
            setattr(self, key, entry)
            self.entries[key] = entry
                
        self.register_button = ttk.Button(self, text="Register Teacher", command=self._register_command)
        self.register_button.pack(pady=15)
        ttk.Button(self, text="Back to Login", command=lambda: controller.show_frame("RoleSelectionPage")).pack()

    def on_show(self, **kwargs):
        """Clear fields on show."""
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def _register_command(self):
        """Registers a new teacher user."""
        data = {key.replace('_entry', ''): entry.get().strip() for key, entry in self.entries.items()}
        
        required_fields = ["name", "subject_name", "subject_code", "phone", "email", "dept_name", "username", "password"]
        if any(not data[f] for f in required_fields):
            MessagePopup(self, "Error", "All fields are required.", msg_type="error")
            return

        if data['username'] in self.controller.users:
            MessagePopup(self, "Error", "Username already taken.", msg_type="error")
            return

        new_user = {
            'role': 'teacher',
            'name': data['name'],
            'subject_name': data['subject_name'],
            'subject_code': data['subject_code'].upper(),
            'phone': data['phone'],
            'email': data['email'],
            'dept_name': data['dept_name'],
            'username': data['username'],
            'password': hash_password(data['password']),
            'created_at': datetime.now().isoformat()
        }
        
        self.controller.users[data['username']] = new_user
        save_data(USERS_FILE, self.controller.users, self.controller)
        
        MessagePopup(self, "Success", "Teacher registered successfully. You can now log in.", msg_type="info")
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.controller.show_frame("RoleSelectionPage")
        
class AdminSignupPage(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.configure(padding="20")
        
        ttk.Label(self, text="Initial Admin Signup", style='Header.TLabel').pack(pady=10)
        
        self.form_frame = ttk.Frame(self)
        self.form_frame.pack(pady=10)
        
        self.initial_admin_exists = any(user.get('role') == 'admin' for user in self.controller.users.values())

        if self.initial_admin_exists:
            ttk.Label(self.form_frame, text="An initial Admin account is already registered.", 
                      foreground='red', font=('Arial', 12, 'bold')).pack(pady=20)
            self.register_button = ttk.Button(self, text="Register Admin", state=tk.DISABLED)
            self.register_button.pack(pady=15)
        else:
            fields = [
                ("Name", "name_entry"), 
                ("Phone Number", "phone_entry"), 
                ("Email ID", "email_entry"), 
                ("Department Name", "dept_name_entry"), # Added Dept Name
                ("Department Code", "dept_code_entry"),
                ("Username", "username_entry"),
                ("Password", "password_entry", True)
            ]
            
            self.entries = {}
            for i, (label_text, key, *is_password) in enumerate(fields):
                ttk.Label(self.form_frame, text=label_text + ":").grid(row=i, column=0, sticky='w', padx=5, pady=5)
                entry = ttk.Entry(self.form_frame, width=40, show="*" if is_password else "")
                entry.grid(row=i, column=1, padx=5, pady=5)
                setattr(self, key, entry)
                self.entries[key] = entry
                
            self.register_button = ttk.Button(self, text="Register Admin", command=self._register_command)
            self.register_button.pack(pady=15)

        ttk.Button(self, text="Back to Login", command=lambda: controller.show_frame("RoleSelectionPage")).pack()

    def _register_command(self):
        """Registers the initial admin user."""
        if self.initial_admin_exists:
            return

        data = {key.replace('_entry', ''): entry.get().strip() for key, entry in self.entries.items()}
        
        required_fields = ["name", "phone", "email", "dept_name", "dept_code", "username", "password"]
        if any(not data[f] for f in required_fields):
            MessagePopup(self, "Error", "All fields are required.", msg_type="error")
            return

        if data['username'] in self.controller.users:
            MessagePopup(self, "Error", "Username already taken.", msg_type="error")
            return

        new_user = {
            'role': 'admin',
            'name': data['name'],
            'phone': data['phone'],
            'email': data['email'],
            'dept_name': data['dept_name'],
            'dept_code': data['dept_code'].upper(),
            'username': data['username'],
            'password': hash_password(data['password']),
            'created_at': datetime.now().isoformat()
        }
        
        self.controller.users[data['username']] = new_user
        save_data(USERS_FILE, self.controller.users, self.controller)
        
        MessagePopup(self, "Success", "Admin registered successfully. You can now log in.", msg_type="info")
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.controller.show_frame("RoleSelectionPage")

# --- Student Dashboard ---

class StudentDashboard(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        
        self.header_frame = ttk.Frame(self)
        self.header_frame.pack(fill='x', padx=10, pady=10)
        
        self.title_label = ttk.Label(self.header_frame, text="Student Dashboard", style='Header.TLabel')
        self.title_label.pack(side=tk.LEFT)
        
        ttk.Button(self.header_frame, text="Logout", command=controller.logout).pack(side=tk.RIGHT)
        
        self.info_frame = ttk.Frame(self)
        self.info_frame.pack(fill='x', padx=20, pady=10)
        
        self.marks_frame = ttk.LabelFrame(self, text="Marks and Results", padding="10")
        self.marks_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.data_labels = {}
        
        # Setup static layout for marks (will be updated dynamically)
        labels = [
            ("Name:", "name"),
            ("USN:", "usn"),
            ("Sem:", "sem"),
            ("Gender:", "gender"),
            ("Phone Number:", "phone"),
            ("Email:", "email"),
            ("Overall Attendance %:", "attendance_perc")
        ]
        
        for i, (label_text, key) in enumerate(labels):
            ttk.Label(self.info_frame, text=label_text, font=('Arial', 10, 'bold')).grid(row=i, column=0, sticky='w', padx=5, pady=2)
            value_label = ttk.Label(self.info_frame, text="...", font=('Arial', 10))
            value_label.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            self.data_labels[key] = value_label
            
        self.marks_tree = ttk.Treeview(self.marks_frame, columns=("Subject", "I1", "I2", "I3", "A1", "A2", "A3", "Lab", "TotalInternal", "External", "TotalMarks"), show='headings')
        self.marks_tree.heading("Subject", text="Subject")
        self.marks_tree.heading("I1", text="I1 (20)")
        self.marks_tree.heading("I2", text="I2 (20)")
        self.marks_tree.heading("I3", text="I3 (20)")
        self.marks_tree.heading("A1", text="A1 (5)")
        self.marks_tree.heading("A2", text="A2 (5)")
        self.marks_tree.heading("A3", text="A3 (5)")
        self.marks_tree.heading("Lab", text="Lab (25)")
        self.marks_tree.heading("TotalInternal", text="Total Internal (50)")
        self.marks_tree.heading("External", text="External (100)")
        self.marks_tree.heading("TotalMarks", text="Total (150)")
        
        self.marks_tree.column("Subject", width=150, anchor=tk.W)
        self.marks_tree.column("TotalInternal", width=100, anchor=tk.CENTER)
        self.marks_tree.column("External", width=70, anchor=tk.CENTER)
        self.marks_tree.column("TotalMarks", width=70, anchor=tk.CENTER)
        
        for col in ["I1", "I2", "I3", "A1", "A2", "A3", "Lab"]:
            self.marks_tree.column(col, width=50, anchor=tk.CENTER)

        self.marks_tree.pack(fill='both', expand=True)

    def refresh_data(self):
        """Fetches and displays the current student's data."""
        if not self.controller.current_user or self.controller.current_user['role'] != 'student':
            return
            
        usn = self.controller.current_user['usn'].upper()
        
        for i in self.marks_tree.get_children():
            self.marks_tree.delete(i)
            
        self.data_labels['usn'].config(text=usn)
        self.data_labels['name'].config(text=self.controller.current_user['name'])
        self.data_labels['email'].config(text=self.controller.current_user['email'])
        self.data_labels['phone'].config(text=self.controller.current_user.get('phone', 'N/A'))
        self.data_labels['sem'].config(text=self.controller.current_user.get('sem', 'N/A'))
        self.data_labels['gender'].config(text=self.controller.current_user.get('gender', 'N/A'))
        
        student_record = [record for record in self.controller.student_data if record['usn'] == usn]
        
        total_classes = 0
        total_attended = 0
        
        for record in student_record:
            subject = record['subject_code']
            
            total_internal, external, total_mark = calculate_marks(record)

            self.marks_tree.insert("", tk.END, values=(
                f"{record['subject_name']} ({subject})",
                record.get('i1', '-'), record.get('i2', '-'), record.get('i3', '-'),
                record.get('a1', '-'), record.get('a2', '-'), record.get('a3', '-'),
                record.get('lab_marks', '-'), total_internal, external, total_mark
            ))
            
            total_classes += record.get('total_classes', 0)
            total_attended += record.get('attended_classes', 0)

        attendance_perc = (total_attended / total_classes * 100) if total_classes > 0 else 0
        self.data_labels['attendance_perc'].config(text=f"{attendance_perc:.2f}%")

# --- Teacher Dashboard ---

class TeacherDashboard(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        
        self.header_frame = ttk.Frame(self)
        self.header_frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(self.header_frame, text="Teacher Dashboard", style='Header.TLabel').pack(side=tk.LEFT)
        ttk.Button(self.header_frame, text="Logout", command=controller.logout).pack(side=tk.RIGHT)
        
        self.subject_label = ttk.Label(self, text="Subject: ---", font=('Arial', 12, 'bold'))
        self.subject_label.pack(pady=(0, 10))

        # Navigation Buttons
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Mark Today's Attendance", 
                   command=lambda: controller.show_frame("TeacherAttendanceEntryPage")).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Enter/Edit All Marks", 
                   command=lambda: controller.show_frame("TeacherMarksEntryPage")).pack(side=tk.LEFT, padx=10)
        
        # --- NEW: Excel Export Button ---
        ttk.Button(button_frame, text="Export Subject Data to Excel", 
                   command=controller.export_to_excel).pack(side=tk.LEFT, padx=10)
        
        # Student List View (Unchanged)
        ttk.Label(self, text="Student List (AI Department)", font=('Arial', 12, 'bold')).pack(pady=10)
        self.student_list_tree = ttk.Treeview(self, columns=("USN", "Name"), show='headings')
        self.student_list_tree.heading("USN", text="USN")
        self.student_list_tree.heading("Name", text="Student Name")
        self.student_list_tree.column("USN", width=150, anchor=tk.CENTER)
        self.student_list_tree.column("Name", width=300, anchor=tk.W)
        self.student_list_tree.pack(fill='both', expand=True, padx=20, pady=10)
        
    def refresh_data(self):
        """Populate the student list and set subject label."""
        if not self.controller.current_user:
            return
            
        subject_name = self.controller.current_user.get('subject_name', 'N/A')
        subject_code = self.controller.current_user.get('subject_code', 'N/A')
        self.subject_label.config(text=f"Subject: {subject_name} ({subject_code})")

        # Clear and repopulate tree
        for i in self.student_list_tree.get_children():
            self.student_list_tree.delete(i)
            
        for usn, name in STUDENT_ROLL.items():
            self.student_list_tree.insert("", tk.END, values=(usn.upper(), name.replace('_', ' ').title()))

# --- Scrollable Frame Widget (for new Teacher pages) ---
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

# --- Teacher Marks Entry Page (REWORKED) ---

class TeacherMarksEntryPage(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.student_mark_widgets = {} # Stores all widgets for each student
        
        ttk.Label(self, text="Mark Entry (All Students)", style='Header.TLabel').pack(pady=10)
        
        # Top button frame
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5, fill='x', padx=10)
        ttk.Button(btn_frame, text="Save All Marks", command=self._save_all_marks).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Load All Data", command=self._load_all_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Back to Dashboard", command=lambda: controller.show_frame("TeacherDashboard")).pack(side=tk.RIGHT, padx=5)

        # Scrollable area for marks
        scroll_frame = ScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.grid_frame = scroll_frame.scrollable_frame

        # --- Create Header Row ---
        headers = ["USN", "Name", "I1", "I2", "I3", "A1", "A2", "A3", "Lab", "External", "Total Int.", "Total"]
        self.mark_keys = ["i1", "i2", "i3", "a1", "a2", "a3", "lab_marks", "external_mark"]
        self.max_vals = {"i1": 20, "i2": 20, "i3": 20, "a1": 5, "a2": 5, "a3": 5, "lab_marks": 25, "external_mark": 100}

        for j, header in enumerate(headers):
            ttk.Label(self.grid_frame, text=header, font=('Arial', 9, 'bold')).grid(row=0, column=j, padx=5, pady=5, sticky='w')

        # --- Create Student Rows ---
        for i, (usn_lower, name) in enumerate(STUDENT_ROLL.items()):
            usn = usn_lower.upper()
            row = i + 1
            widgets = {} # To store this student's widgets
            
            # USN and Name
            ttk.Label(self.grid_frame, text=usn).grid(row=row, column=0, padx=5, sticky='w')
            ttk.Label(self.grid_frame, text=name.replace('_', ' ').title()).grid(row=row, column=1, padx=5, sticky='w')
            
            col_idx = 2
            # Mark Entry Boxes
            for key in self.mark_keys:
                entry = ttk.Entry(self.grid_frame, width=5)
                entry.grid(row=row, column=col_idx, padx=2)
                widgets[key] = entry
                col_idx += 1
            
            # Calculated Labels
            internal_label = ttk.Label(self.grid_frame, text="--", width=5)
            internal_label.grid(row=row, column=col_idx, padx=5)
            widgets['total_internal_label'] = internal_label
            col_idx += 1
            
            total_label = ttk.Label(self.grid_frame, text="--", width=5)
            total_label.grid(row=row, column=col_idx, padx=5)
            widgets['total_mark_label'] = total_label
            
            self.student_mark_widgets[usn] = widgets

    def on_show(self, **kwargs):
        """Load data when page is shown."""
        self._load_all_data()

    def _load_all_data(self):
        """Loads all marks for all students for this teacher's subject."""
        teacher_subject_code = self.controller.current_user['subject_code']
        
        for usn, widgets in self.student_mark_widgets.items():
            # Find the specific record
            record = next((r for r in self.controller.student_data 
                           if r['usn'] == usn and r['subject_code'] == teacher_subject_code), None)
            
            # Clear all entry fields first
            for key in self.mark_keys:
                widgets[key].delete(0, tk.END)
            
            widgets['total_internal_label'].config(text="--")
            widgets['total_mark_label'].config(text="--")

            if record:
                # Populate entry fields
                for key in self.mark_keys:
                    widgets[key].insert(0, str(record.get(key, 0)))
                
                # Calculate and set labels
                total_internal, _, total_mark = calculate_marks(record)
                widgets['total_internal_label'].config(text=str(total_internal))
                widgets['total_mark_label'].config(text=str(total_mark))
        
        MessagePopup(self, "Data Loaded", "All student data for this subject has been loaded.", msg_type="info")


    def _save_all_marks(self):
        """Saves all marks for all students for this subject."""
        teacher_subject_code = self.controller.current_user['subject_code']
        teacher_subject_name = self.controller.current_user['subject_name']
        
        try:
            for usn, widgets in self.student_mark_widgets.items():
                marks_data = {}
                # Read from entry boxes
                for key in self.mark_keys:
                    val_str = widgets[key].get().strip()
                    val = int(val_str or 0)
                    
                    if val < 0 or val > self.max_vals[key]:
                        raise ValueError(f"USN {usn}: {key} must be between 0 and {self.max_vals[key]}.")
                    
                    marks_data[key] = val

                # Find existing record or create new
                record = next((r for r in self.controller.student_data 
                               if r['usn'] == usn and r['subject_code'] == teacher_subject_code), None)
                
                if not record:
                    record = {'usn': usn, 'subject_code': teacher_subject_code}
                    self.controller.student_data.append(record)
                
                # Update record
                record.update(marks_data)
                record['subject_name'] = teacher_subject_name # Ensure name is set
                
                # Recalculate and update labels
                total_internal, _, total_mark = calculate_marks(record)
                widgets['total_internal_label'].config(text=str(total_internal))
                widgets['total_mark_label'].config(text=str(total_mark))
            
            # Save to file
            save_data(STUDENT_DATA_FILE, self.controller.student_data, self.controller)
            MessagePopup(self, "Success", "All marks saved successfully.", msg_type="info")
            
        except ValueError as e:
            MessagePopup(self, "Validation Error", str(e), msg_type="error")
        except Exception as e:
            MessagePopup(self, "Error", f"An unexpected error occurred: {e}", msg_type="error")


# --- Teacher Attendance Entry Page (REWORKED) ---

class TeacherAttendanceEntryPage(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.attendance_vars = {} # Stores tk.StringVar for each student
        
        ttk.Label(self, text="Mark Today's Attendance", style='Header.TLabel').pack(pady=10)
        
        # Top button frame
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5, fill='x', padx=10)
        ttk.Button(btn_frame, text="Save Attendance", command=self._save_attendance).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Back to Dashboard", command=lambda: controller.show_frame("TeacherDashboard")).pack(side=tk.RIGHT, padx=5)

        # Scrollable area for attendance
        scroll_frame = ScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.grid_frame = scroll_frame.scrollable_frame

        # --- Create Header Row ---
        headers = ["USN", "Name", "Status"]
        for j, header in enumerate(headers):
            ttk.Label(self.grid_frame, text=header, font=('Arial', 9, 'bold')).grid(row=0, column=j, padx=5, pady=5, sticky='w')

        # --- Create Student Rows ---
        for i, (usn_lower, name) in enumerate(STUDENT_ROLL.items()):
            usn = usn_lower.upper()
            row = i + 1
            
            # USN and Name
            ttk.Label(self.grid_frame, text=usn).grid(row=row, column=0, padx=5, sticky='w')
            ttk.Label(self.grid_frame, text=name.replace('_', ' ').title()).grid(row=row, column=1, padx=5, sticky='w')
            
            # Radio Buttons
            var = tk.StringVar(value="N/A") # Default to N/A
            self.attendance_vars[usn] = var
            
            radio_frame = ttk.Frame(self.grid_frame)
            radio_frame.grid(row=row, column=2, padx=5)
            
            ttk.Radiobutton(radio_frame, text="Present", variable=var, value="Present").pack(side=tk.LEFT)
            ttk.Radiobutton(radio_frame, text="Absent", variable=var, value="Absent").pack(side=tk.LEFT)
            ttk.Radiobutton(radio_frame, text="N/A", variable=var, value="N/A").pack(side=tk.LEFT)

    def on_show(self, **kwargs):
        """Reset all radio buttons to N/A when page is shown."""
        for var in self.attendance_vars.values():
            var.set("N/A")

    def _save_attendance(self):
        """Increments attendance counts based on radio button selection."""
        teacher_subject_code = self.controller.current_user['subject_code']
        teacher_subject_name = self.controller.current_user['subject_name']
        
        try:
            for usn, var in self.attendance_vars.items():
                status = var.get()
                
                if status == "N/A":
                    continue # Skip students not marked
                
                # Find existing record or create new
                record = next((r for r in self.controller.student_data 
                               if r['usn'] == usn and r['subject_code'] == teacher_subject_code), None)
                
                if not record:
                    record = {'usn': usn, 'subject_code': teacher_subject_code, 'subject_name': teacher_subject_name,
                              'total_classes': 0, 'attended_classes': 0}
                    self.controller.student_data.append(record)
                
                # Increment counts
                record['total_classes'] = record.get('total_classes', 0) + 1
                if status == "Present":
                    record['attended_classes'] = record.get('attended_classes', 0) + 1
            
            # Save to file
            save_data(STUDENT_DATA_FILE, self.controller.student_data, self.controller)
            MessagePopup(self, "Success", "Today's attendance has been saved.", msg_type="info")
            
            # Reset page
            self.on_show()

        except Exception as e:
            MessagePopup(self, "Error", f"An unexpected error occurred: {e}", msg_type="error")

# --- Admin Dashboard ---

class AdminDashboard(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        
        self.header_frame = ttk.Frame(self)
        self.header_frame.pack(fill='x', padx=10, pady=10)
        ttk.Label(self.header_frame, text="Admin Dashboard: All Data", style='Header.TLabel').pack(side=tk.LEFT)
        ttk.Button(self.header_frame, text="Logout", command=controller.logout).pack(side=tk.RIGHT)
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5, fill='x', padx=10)
        ttk.Button(btn_frame, text="Refresh Data View", command=self.refresh_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Export All Data to Excel", 
                   command=controller.export_to_excel).pack(side=tk.LEFT, padx=5)

        # Treeview to display all student data
        self.data_tree = ttk.Treeview(self, columns=(
            "USN", "Name", "Department", "Subject", "Total Internal", "Attendance %", "External Mark", "Total Mark"
        ), show='headings')
        
        self.data_tree.heading("USN", text="USN")
        self.data_tree.heading("Name", text="Student Name")
        self.data_tree.heading("Department", text="Department")
        self.data_tree.heading("Subject", text="Subject Code")
        self.data_tree.heading("Total Internal", text="Total Internal (50)")
        self.data_tree.heading("Attendance %", text="Attendance %")
        self.data_tree.heading("External Mark", text="External Mark (100)")
        self.data_tree.heading("Total Mark", text="Total Mark (150)")

        self.data_tree.column("USN", width=100, anchor=tk.CENTER)
        self.data_tree.column("Name", width=150, anchor=tk.W)
        self.data_tree.column("Department", width=120, anchor=tk.W)
        self.data_tree.column("Subject", width=80, anchor=tk.CENTER)
        self.data_tree.column("Total Internal", width=100, anchor=tk.CENTER)
        self.data_tree.column("Attendance %", width=100, anchor=tk.CENTER)
        self.data_tree.column("External Mark", width=100, anchor=tk.CENTER)
        self.data_tree.column("Total Mark", width=100, anchor=tk.CENTER)

        self.data_tree.pack(fill='both', expand=True, padx=20, pady=10)
        
    def refresh_data(self):
        """Populates the Treeview with all student mark and attendance data."""
        for i in self.data_tree.get_children():
            self.data_tree.delete(i)
        
        # Create dictionaries to hold student info indexed by USN
        student_info = {user['usn'].upper(): {'name': user['name'], 'dept': user.get('dept_name', 'N/A')}
                         for user in self.controller.users.values() if user['role'] == 'student'}
        
        for record in self.controller.student_data:
            usn = record['usn']
            
            # Recalculate marks and attendance percentage
            total_internal, external_mark, total_mark = calculate_marks(record)
            
            total_classes = record.get('total_classes', 0)
            attended_classes = record.get('attended_classes', 0)
            attendance_perc = (attended_classes / total_classes * 100) if total_classes > 0 else 0
            
            name = student_info.get(usn, {}).get('name', 'N/A')
            dept = student_info.get(usn, {}).get('dept', 'N/A')
            
            self.data_tree.insert("", tk.END, values=(
                usn, name, dept, record['subject_code'], total_internal, f"{attendance_perc:.2f}", 
                external_mark, total_mark
            ))

# --- Admin Creation Page (Access via 'Iceberg Menu') ---

class AdminCreationPage(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        
        ttk.Label(self, text="Create New Admin Account", style='Header.TLabel').pack(pady=10)
        
        form_frame = ttk.Frame(self)
        form_frame.pack(pady=10)
        
        fields = [
            ("Admin Name", "name_entry"), 
            ("Phone Number", "phone_entry"), 
            ("Email ID", "email_entry"), 
            ("Department Name", "dept_name_entry"), # Added Dept Name
            ("Department Code", "dept_code_entry"),
            ("Username", "username_entry"),
            ("Password", "password_entry", True),
            ("App Code", "app_code_entry", True)
        ]
        
        self.entries = {}
        for i, (label_text, key, *is_password) in enumerate(fields):
            ttk.Label(form_frame, text=label_text + ":").grid(row=i, column=0, sticky='w', padx=5, pady=5)
            entry = ttk.Entry(form_frame, width=40, show="*" if is_password and key == 'password_entry' else "")
            entry.grid(row=i, column=1, padx=5, pady=5)
            setattr(self, key, entry)
            self.entries[key] = entry
            
        ttk.Button(self, text="Create Admin", command=self._create_admin).pack(pady=15)
        ttk.Button(self, text="Back to Dashboard", command=lambda: controller.show_frame("AdminDashboard")).pack()
        
    def _create_admin(self):
        """Creates a new admin account after verifying the App Code."""
        data = {key.replace('_entry', ''): entry.get().strip() for key, entry in self.entries.items()}
        
        required_fields = ["name", "phone", "email", "dept_name", "dept_code", "username", "password", "app_code"]
        if any(not data[f] for f in required_fields):
            MessagePopup(self, "Error", "All fields are required.", msg_type="error")
            return

        if data['app_code'] != APP_CODE:
            MessagePopup(self, "Error", "Invalid App Code. Admin creation failed.", msg_type="error")
            return

        if data['username'] in self.controller.users:
            MessagePopup(self, "Error", "Username already taken.", msg_type="error")
            return

        new_admin = {
            'role': 'admin',
            'name': data['name'],
            'phone': data['phone'],
            'email': data['email'],
            'dept_name': data['dept_name'],
            'dept_code': data['dept_code'].upper(),
            'username': data['username'],
            'password': hash_password(data['password']),
            'created_at': datetime.now().isoformat()
        }
        
        self.controller.users[data['username']] = new_admin
        save_data(USERS_FILE, self.controller.users, self.controller)
        
        MessagePopup(self, "Success", "New Admin account created successfully.", msg_type="info")
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.controller.show_frame("AdminDashboard")


if __name__ == "__main__":
    app = SISApp()
    app.mainloop()

