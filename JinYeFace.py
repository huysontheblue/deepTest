import os
import json
import cv2
import numpy as np
import face_recognition
import dlib
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import datetime
import pandas as pd
import threading
import time
import shutil
import sqlite3
from functools import partial


class FaceRecognitionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Nhận Diện Khuôn Mặt人脸识别")
        self.root.geometry("1000x600")
        self.root.resizable(False, False)

        # Thiết lập đường dẫn và thư mục
        self.network_paths = {
            "base_dir": r"\\10.2.100.200\Jinyebaoan",
            "background_path": r"\\10.2.100.200\Jinyebaoan\image.jpg",
            "icon_path": r"\\10.2.100.200\Jinyebaoan\icon.ico",
            "database_dir": r"\\10.2.100.200\Jinyebaoan\database",
            "dlib_model_path": r"\\10.2.100.200\Jinyebaoan\dlib_face_recognition_resnet_model_v1.dat",
            "shape_predictor_path": r"\\10.2.100.200\Jinyebaoan\shape_predictor_68_face_landmarks.dat"
        }

        # Đảm bảo thư mục cơ sở tồn tại
        os.makedirs(self.network_paths["base_dir"], exist_ok=True)
        os.makedirs(self.network_paths["database_dir"], exist_ok=True)

        # Đường dẫn đến file SQLite
        self.db_path = os.path.join(self.network_paths["database_dir"], "face_recognition.db")

        # Khởi tạo biến cho camera
        self.cap = None
        self.camera_index = 0
        self.is_capturing = False
        self.verification_thread = None

        # Hiển thị màn hình đăng nhập trước khi vào ứng dụng chính
        self.show_login_screen()

    def create_database(self):
        """Tạo cơ sở dữ liệu SQLite nếu chưa tồn tại"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Tạo bảng users
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                is_admin INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Tạo bảng employees
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                employee_id TEXT UNIQUE NOT NULL,
                department TEXT NOT NULL,
                permission TEXT NOT NULL,
                image_folder TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Tạo bảng face_encodings
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS face_encodings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                encoding BLOB NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees (employee_id)
            )
        """)

        # Tạo bảng work_areas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS work_areas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                area TEXT NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees (employee_id)
            )
        """)

        # Tạo bảng history
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT NOT NULL,
                name TEXT NOT NULL,
                department TEXT NOT NULL,
                permission TEXT NOT NULL,
                area TEXT NOT NULL,
                is_authorized INTEGER NOT NULL,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (employee_id) REFERENCES employees (employee_id)
            )
        """)

        conn.commit()
        conn.close()

    def show_login_screen(self):
        # Xóa tất cả widget hiện tại trên root
        for widget in self.root.winfo_children():
            widget.destroy()

        # Tạo cơ sở dữ liệu nếu chưa tồn tại
        self.create_database()

        # Kiểm tra và tạo tài khoản admin mặc định nếu chưa có người dùng
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        if cursor.fetchone()[0] == 0:
            # Tạo tài khoản admin mặc định
            cursor.execute("""INSERT INTO users (username, password, is_admin) VALUES (?, ?, ?)""", ("123", "123", 1))
            conn.commit()
        conn.close()

        # Thêm hình nền cho toàn bộ root (giao diện đăng nhập)
        if os.path.exists(self.network_paths["background_path"]):
            try:
                bg_image = Image.open(self.network_paths["background_path"])
                bg_image = bg_image.resize((1000, 600), Image.Resampling.LANCZOS)
                self.root_background_image = ImageTk.PhotoImage(bg_image)
                self.root_background_label = tk.Label(self.root, image=self.root_background_image)
                self.root_background_label.place(x=0, y=0, relwidth=1, relheight=1)
            except Exception as e:
                print(f"Lỗi khi tải ảnh nền cho root: {e}")
                self.root.configure(bg="#f0f0f0")  # Màu nền mặc định nếu lỗi
        else:
            self.root.configure(bg="#f0f0f0")

        # Tạo frame đăng nhập
        login_frame = tk.Frame(self.root, bg="#f0f0f0")
        login_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER, width=400, height=450)

        # Tiêu đề đăng nhập
        login_title = tk.Label(login_frame, text="ĐĂNG NHẬP HỆ THỐNG\n系统登录", font=("Arial", 18, "bold"),
                               bg="#f0f0f0")
        login_title.pack(pady=20)

        # Frame chứa các trường nhập liệu
        input_frame = tk.Frame(login_frame, bg="#f0f0f0")
        input_frame.pack(fill=tk.X, padx=30)

        # Tên đăng nhập
        tk.Label(input_frame, text="Tên đăng nhập\用户名:", font=("Arial", 12), bg="#f0f0f0", anchor="w").pack(fill=tk.X)
        self.username_entry = tk.Entry(input_frame, font=("Arial", 12))
        self.username_entry.pack(fill=tk.X, pady=(0, 10))

        # Mật khẩu
        tk.Label(input_frame, text="Mật khẩu\密码:", font=("Arial", 12), bg="#f0f0f0", anchor="w").pack(fill=tk.X)
        self.password_entry = tk.Entry(input_frame, font=("Arial", 12), show="*")
        self.password_entry.pack(fill=tk.X, pady=(0, 10))

        # Khu vực làm việc
        tk.Label(input_frame, text="Khu vực làm việc\n工作区域:",font=("Arial", 12), bg="#f0f0f0", anchor="w").pack(fill=tk.X)
        self.work_area_combobox = ttk.Combobox(input_frame, font=("Arial", 12),values=["A", "B", "C", "D", "E", "Thí nghiệm"])
        self.work_area_combobox.pack(fill=tk.X, pady=(0, 20))
        self.work_area_combobox.current(0)  # Mặc định chọn khu vực A

        # Nút đăng nhập - đặt trong input_frame để hiển thị đúng
        login_button = tk.Button(input_frame, text="ĐĂNG NHẬP\n登录", font=("Arial", 14, "bold"), bg="#2980b9",fg="white", relief=tk.RAISED, command=self.validate_login)
        login_button.pack(fill=tk.X, pady=10)

        # Thêm dòng chữ "Programmed IT Việt Nam PMP - @2025" ở dưới cùng
        powered_label = tk.Label(login_frame, text="Programmed IT Việt Nam PMP - @2025", font=("Arial", 10),bg="#f0f0f0", fg="#666666")
        powered_label.pack(side=tk.BOTTOM, pady=10)
        #hubert_label = tk.Label(login_frame, text="Hubert_He", font=("Arial", 4), bg="#f0f0f0", fg="#666666")
        #hubert_label.pack(side=tk.BOTTOM, pady=5)
        # Focus vào ô tên đăng nhập
        self.username_entry.focus()

        # Bắt sự kiện Enter cho tất cả các trường nhập liệu
        self.username_entry.bind("<Return>", lambda event: self.password_entry.focus())
        self.password_entry.bind("<Return>", lambda event: self.work_area_combobox.focus())
        self.work_area_combobox.bind("<Return>", lambda event: self.validate_login())

        # Bắt sự kiện khi người dùng chọn combobox
        self.work_area_combobox.bind("<<ComboboxSelected>>", lambda event: login_button.focus())

    def validate_login(self):
        """Xác thực đăng nhập"""
        username = self.username_entry.get()
        password = self.password_entry.get()
        work_area = self.work_area_combobox.get()

        if not username or not password:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!\n请填写完整信息!")
            return

        # Xác thực tài khoản từ cơ sở dữ liệu
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            # Kiểm tra cấu trúc bảng users trước khi truy vấn
            cursor.execute("PRAGMA table_info(users)")
            columns = [info[1] for info in cursor.fetchall()]

            if 'password' in columns:
                # Nếu có cột password, sử dụng truy vấn ban đầu
                cursor.execute("""
                    SELECT id, is_admin FROM users 
                    WHERE username = ? AND password = ?
                """, (username, password))
            else:
                # Nếu không có cột password, chỉ xác thực bằng username
                # Đây là giải pháp tạm thời cho đến khi cấu trúc bảng được cập nhật
                cursor.execute("""
                    SELECT id, is_admin FROM users 
                    WHERE username = ?
                """, (username,))
                print("Cảnh báo: Đang đăng nhập không kiểm tra mật khẩu do cấu trúc CSDL thiếu cột 'password'")

            user_data = cursor.fetchone()
        except sqlite3.OperationalError as e:
            # Xử lý trường hợp lỗi cấu trúc bảng
            print(f"Lỗi khi truy vấn CSDL: {e}")
            messagebox.showerror("Lỗi CSDL", f"Có lỗi với cơ sở dữ liệu: {e}\nVui lòng khởi động lại ứng dụng!")
            conn.close()
            return
        finally:
            conn.close()

        if user_data:
            self.current_user_id = user_data[0]
            self.is_admin = user_data[1] == 1
            self.current_work_area = work_area

            # Lưu thông tin đăng nhập
            # messagebox.showinfo("Thành công", f"Đăng nhập thành công!\n登录成功! \nKhu vực làm việc: {work_area}")

            # Khởi tạo các thành phần cần thiết
            self.initialize_app()
        else:
            messagebox.showerror("Lỗi", "Tên đăng nhập hoặc mật khẩu không đúng!\n用户名或密码错误!")

    def initialize_app(self):
        """Khởi tạo các thành phần chính của ứng dụng sau khi đăng nhập thành công"""
        # Xóa tất cả widget hiện tại
        for widget in self.root.winfo_children():
            widget.destroy()

        # Tạo các thư mục cần thiết
        self.create_directories()

        # Thiết lập icon cho ứng dụng
        if os.path.exists(self.network_paths["icon_path"]):
            self.root.iconbitmap(self.network_paths["icon_path"])

        # Thiết lập ảnh nền
        if os.path.exists(self.network_paths["background_path"]):
            try:
                bg_image = Image.open(self.network_paths["background_path"])
                bg_image = bg_image.resize((1000, 600))
                self.background_image = ImageTk.PhotoImage(bg_image)

                self.background_label = tk.Label(self.root, image=self.background_image)
                self.background_label.place(x=0, y=0, relwidth=1, relheight=1)
            except Exception as e:
                print(f"Lỗi khi tải ảnh nền: {e}")

        # Load face detector
        self.detector = dlib.get_frontal_face_detector()
        if os.path.exists(self.network_paths["shape_predictor_path"]):
            self.predictor = dlib.shape_predictor(self.network_paths["shape_predictor_path"])

        # Tạo giao diện
        self.create_top_menu()
        self.create_buttons()

    def create_directories(self):
        """Tạo các thư mục cần thiết"""
        # Tạo thư mục lưu ảnh nhân viên
        self.employee_images_dir = os.path.join(self.network_paths["database_dir"], "employee_images")
        os.makedirs(self.employee_images_dir, exist_ok=True)

    def create_top_menu(self):
        """Tạo thanh menu phía trên"""
        # Menu bar
        menu_frame = tk.Frame(self.root, bg="#0078D7", height=40)
        menu_frame.pack(fill=tk.X)

        # Hiển thị thông tin người dùng đăng nhập
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT username FROM users WHERE id = ?", (self.current_user_id,))
        username = cursor.fetchone()[0]
        conn.close()

        # Label hiển thị người dùng
        user_label = tk.Label(
            menu_frame,
            text=f"Người dùng: {username} | Khu vực: {self.current_work_area}",
            font=("Arial", 10),
            bg="#0078D7",
            fg="white"
        )
        user_label.pack(side=tk.LEFT, padx=10, pady=8)

        # Menu buttons
        menu_items = [
            ("Quản lý người dùng\n用户管理", self.manage_users, "#1a5276"),
            ("Danh sách nhân viên\n员工名单", self.show_employee_list, "#1f618d"),
            ("Lịch sử\n历史", self.show_history, "#2874a6"),
            ("Xuất Excel\n导出Excel", self.export_options, "#2e86c1"),
            ("Đăng xuất\n登出", self.logout, "#e74c3c")
        ]

        # Chỉ hiển thị menu quản lý người dùng nếu là admin
        if not self.is_admin:
            menu_items = menu_items[4:]

        for text, command, bg_color in menu_items:
            button = tk.Button(
                menu_frame,
                text=text,
                font=("Arial", 10, "bold"),
                bg=bg_color,
                fg="white",
                relief=tk.FLAT,
                command=command,
                padx=10
            )
            button.pack(side=tk.LEFT, padx=2, pady=4)

    def on_camera_selected(self, event):
        """Xử lý khi chọn camera từ combobox"""
        selected = self.camera_combobox.current()
        if 0 <= selected < len(self.available_cameras):
            self.camera_index = self.available_cameras[selected]

    # quyền hạn 2 nút đăng ký xác thực
    def create_buttons(self):
        # Tạo Frame bao ngoài, đặt ở dưới cùng giao diện
        button_frame = tk.Frame(self.root, bg="", height=100)
        button_frame.place(x=0, y=500, width=1000, height=100)

        # Nút Đăng Ký
        register_button_state = tk.NORMAL if self.is_admin else tk.DISABLED
        register_button_bg = "#2E86C1" if self.is_admin else "#A9A9A9"  # Màu xám khi bị khóa
        register_button = tk.Button(
            button_frame,
            text="ĐĂNG KÝ KHUÔN MẶT\n人脸注册",
            font=("Arial", 14, "bold"),
            bg=register_button_bg,
            fg="white",
            relief=tk.RAISED,
            state=register_button_state,
            command=self.open_registration_form
        )
        register_button.place(x=0, y=0, width=500, height=120)

        # Nút Xác Thực
        verify_button = tk.Button(
            button_frame,
            text="XÁC THỰC KHUÔN MẶT\n人脸验证",
            font=("Arial", 14, "bold"),
            bg="#3498DB",
            fg="white",
            relief=tk.RAISED,
            command=self.open_face_verification
        )
        verify_button.place(x=500, y=0, width=500, height=120)

    def stop_camera_test(self, window):
        """Dừng test camera và đóng cửa sổ"""
        self.is_testing = False
        time.sleep(0.2)  # Đợi thread kết thúc
        window.destroy()

    def logout(self):
        """Đăng xuất khỏi hệ thống"""
        # Hỏi xác nhận
        confirm = messagebox.askyesno("Xác nhận", "Bạn có chắc muốn đăng xuất?\n确定要登出吗?")
        if confirm:
            # Dừng tất cả các thread đang chạy
            if self.is_capturing:
                self.is_capturing = False

            if self.verification_thread and self.verification_thread.is_alive():
                time.sleep(0.5)  # Đợi thread kết thúc

            # Đóng tất cả cửa sổ con
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Toplevel):
                    widget.destroy()

            # Hiển thị lại màn hình đăng nhập
            self.show_login_screen()

    def manage_users(self):
        """Quản lý người dùng (chỉ admin)"""
        if not self.is_admin:
            messagebox.showwarning("Cảnh báo", "Bạn không có quyền truy cập chức năng này!\n您无权访问此功能!")
            return

        # Tạo cửa sổ quản lý người dùng
        user_window = tk.Toplevel(self.root)
        user_window.title("Quản Lý Người Dùng用户管理")
        user_window.geometry("600x500")

        # Frame tiêu đề
        title_frame = tk.Frame(user_window, bg="#0078D7")
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="QUẢN LÝ NGƯỜI DÙNG\n用户管理",
            font=("Arial", 14, "bold"),
            bg="#0078D7",
            fg="white",
            pady=10
        ).pack()

        # Frame chứa nút thêm người dùng
        add_frame = tk.Frame(user_window)
        add_frame.pack(fill=tk.X, pady=10)

        add_button = tk.Button(
            add_frame,
            text="Thêm người dùng\n添加用户",
            font=("Arial", 12),
            bg="#27ae60",
            fg="white",
            command=lambda: self.add_user_form(user_window)
        )
        add_button.pack(padx=10, pady=5)

        # Treeview hiển thị danh sách người dùng
        columns = ("id", "username", "is_admin", "created_at")
        self.user_treeview = ttk.Treeview(user_window, columns=columns, show="headings")

        # Thiết lập tiêu đề cột
        self.user_treeview.heading("id", text="ID")
        self.user_treeview.heading("username", text="Tên đăng nhập\n用户名")
        self.user_treeview.heading("is_admin", text="Quyền admin\n管理员权限")
        self.user_treeview.heading("created_at", text="Ngày tạo\n创建日期")

        # Thiết lập chiều rộng cột
        self.user_treeview.column("id", width=50)
        self.user_treeview.column("username", width=150)
        self.user_treeview.column("is_admin", width=150)
        self.user_treeview.column("created_at", width=200)

        # Thêm dữ liệu vào treeview
        self.populate_user_treeview()

        # Scrollbar
        scrollbar = ttk.Scrollbar(user_window, orient=tk.VERTICAL, command=self.user_treeview.yview)
        self.user_treeview.configure(yscrollcommand=scrollbar.set)

        # Đặt treeview và scrollbar
        treeview_frame = tk.Frame(user_window)
        treeview_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        self.user_treeview.pack(expand=True, fill=tk.BOTH, side=tk.LEFT)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Frame chứa các nút
        button_frame = tk.Frame(user_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        # Nút Xóa người dùng
        delete_button = tk.Button(
            button_frame,
            text="Xóa\n删除",
            font=("Arial", 12),
            bg="#e74c3c",
            fg="white",
            command=lambda: self.delete_user()
        )
        delete_button.pack(side=tk.LEFT, padx=5)

        # Nút Đổi mật khẩu
        change_pwd_button = tk.Button(
            button_frame,
            text="Đổi mật khẩu\n更改密码",
            font=("Arial", 12),
            bg="#f39c12",
            fg="white",
            command=lambda: self.change_user_password()
        )
        change_pwd_button.pack(side=tk.LEFT, padx=5)

    def populate_user_treeview(self):
        """Cập nhật danh sách người dùng vào treeview"""
        # Xóa dữ liệu cũ
        for item in self.user_treeview.get_children():
            self.user_treeview.delete(item)

        # Lấy dữ liệu từ SQLite
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, is_admin, created_at FROM users ORDER BY id")

        for row in cursor.fetchall():
            is_admin = "Admin" if row[2] == 1 else "Người dùng thường"
            self.user_treeview.insert("", tk.END, values=(row[0], row[1], is_admin, row[3]))

        conn.close()

    def add_user_form(self, parent_window):
        """Hiển thị form thêm người dùng mới"""
        # Tạo cửa sổ thêm người dùng
        add_window = tk.Toplevel(parent_window)
        add_window.title("Thêm Người Dùng添加用户")
        add_window.geometry("400x320")  # THAY ĐỔI 1: Tăng chiều cao từ 250 lên 320

        # Frame chứa tiêu đề
        # THAY ĐỔI 2: Thêm frame tiêu đề để giao diện nhất quán
        title_frame = tk.Frame(add_window, bg="#27ae60")
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="THÊM NGƯỜI DÙNG\n添加用户",
            font=("Arial", 14, "bold"),
            bg="#27ae60",
            fg="white",
            pady=10
        ).pack()

        # Frame chứa các trường nhập liệu
        input_frame = tk.Frame(add_window)
        input_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)  # THAY ĐỔI 3: Giảm pady từ 20 xuống 10

        # Tên đăng nhập
        tk.Label(input_frame, text="Tên đăng nhập\n用户名:", font=("Arial", 12), anchor="w").grid(row=0, column=0,sticky="w", pady=5)
        username_entry = tk.Entry(input_frame, font=("Arial", 12), width=25)
        username_entry.grid(row=0, column=1, sticky="w", pady=5)

        # Mật khẩu
        tk.Label(input_frame, text="Mật khẩu\n密码:", font=("Arial", 12), anchor="w").grid(row=1, column=0, sticky="w",pady=5)
        password_entry = tk.Entry(input_frame, font=("Arial", 12), width=25, show="*")
        password_entry.grid(row=1, column=1, sticky="w", pady=5)

        # Xác nhận mật khẩu
        tk.Label(input_frame, text="Xác nhận mật khẩu\n确认密码:", font=("Arial", 12), anchor="w").grid(row=2, column=0,sticky="w",pady=5)
        confirm_password_entry = tk.Entry(input_frame, font=("Arial", 12), width=25, show="*")
        confirm_password_entry.grid(row=2, column=1, sticky="w", pady=5)

        # Loại tài khoản
        tk.Label(input_frame, text="Loại tài khoản\n账户类型:", font=("Arial", 12), anchor="w").grid(row=3, column=0,sticky="w", pady=5)
        admin_var = tk.IntVar()
        admin_checkbox = tk.Checkbutton(input_frame, text="Tài khoản admin\n管理员账户", variable=admin_var,font=("Arial", 12))
        admin_checkbox.grid(row=3, column=1, sticky="w", pady=5)

        # THAY ĐỔI 4: Tạo frame riêng cho nút lưu thay vì đặt trong input_frame
        button_frame = tk.Frame(add_window)
        button_frame.pack(fill=tk.X, padx=20, pady=10)

        # Nút Thêm - THAY ĐỔI 5: Chuyển từ grid sang pack và đặt trong button_frame
        add_button = tk.Button(
            button_frame,
            text="THÊM\n添加",
            font=("Arial", 12, "bold"),
            bg="#27ae60",
            fg="white",
            command=lambda: self.save_new_user(add_window, username_entry.get(), password_entry.get(),confirm_password_entry.get(), admin_var.get())
        )
        add_button.pack(fill=tk.X, pady=5)  # THAY ĐỔI 6: Sử dụng pack() với fill=tk.X

        # Focus vào ô tên đăng nhập
        username_entry.focus()

        # THAY ĐỔI 7: Thêm bắt sự kiện Enter để dễ dàng điền form
        username_entry.bind("<Return>", lambda event: password_entry.focus())
        password_entry.bind("<Return>", lambda event: confirm_password_entry.focus())
        confirm_password_entry.bind("<Return>", lambda event: admin_checkbox.focus())
        admin_checkbox.bind("<Return>", lambda event: add_button.invoke())  # Kích hoạt nút khi Enter

    def save_new_user(self, window, username, password, confirm_password, is_admin):
        """Lưu người dùng mới vào cơ sở dữ liệu"""
        # Kiểm tra thông tin nhập
        if not username or not password:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!\n请填写完整信息!")
            return

        if password != confirm_password:
            messagebox.showerror("Lỗi", "Mật khẩu xác nhận không khớp!\n确认密码不匹配!")
            return

        # Kiểm tra tên đăng nhập đã tồn tại chưa
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM users WHERE username = ?", (username,))
        if cursor.fetchone()[0] > 0:
            conn.close()
            messagebox.showerror("Lỗi", f"Tên đăng nhập '{username}' đã tồn tại!\n用户名已存在!")
            return

        # Thêm người dùng mới
        try:
            cursor.execute("""
                INSERT INTO users (username, password, is_admin) 
                VALUES (?, ?, ?)
            """, (username, password, 1 if is_admin else 0))

            conn.commit()
            conn.close()

            # Cập nhật lại treeview
            self.populate_user_treeview()

            # Đóng cửa sổ
            window.destroy()

            messagebox.showinfo("Thành công", f"Đã thêm người dùng '{username}' thành công!\n已成功添加用户!")
        except Exception as e:
            conn.close()
            messagebox.showerror("Lỗi", f"Không thể thêm người dùng: {str(e)}")

    def delete_user(self):
        """Xóa người dùng"""
        # Lấy người dùng được chọn
        selected_item = self.user_treeview.selection()

        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn người dùng cần xóa!\n请选择要删除的用户!")
            return

        # Lấy thông tin người dùng
        values = self.user_treeview.item(selected_item[0], "values")
        user_id = values[0]
        username = values[1]

        # Không cho phép xóa tài khoản đang sử dụng
        if int(user_id) == self.current_user_id:
            messagebox.showwarning("Cảnh báo", "Không thể xóa tài khoản đang sử dụng!\n无法删除当前使用的账户!")
            return

        # Xác nhận xóa
        confirm = messagebox.askyesno("Xác nhận",
                                      f"Bạn có chắc muốn xóa người dùng '{username}'?\n确定要删除用户 '{username}' 吗?")

        if confirm:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            # Xóa người dùng
            try:
                cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
                conn.commit()

                # Cập nhật lại treeview
                self.user_treeview.delete(selected_item[0])

                messagebox.showinfo("Thành công", f"Đã xóa người dùng '{username}' thành công!\n已成功删除用户!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xóa người dùng: {str(e)}")
            finally:
                conn.close()

    def change_user_password(self):
        """Đổi mật khẩu cho người dùng"""
        # Lấy người dùng được chọn
        selected_item = self.user_treeview.selection()

        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn người dùng cần đổi mật khẩu!\n请选择要更改密码的用户!")
            return

        # Lấy thông tin người dùng
        values = self.user_treeview.item(selected_item[0], "values")
        user_id = values[0]
        username = values[1]

        # Tạo cửa sổ đổi mật khẩu
        pwd_window = tk.Toplevel(self.root)
        pwd_window.title(f"Đổi Mật Khẩu - {username}更改密码")
        pwd_window.geometry("400x200")

        # Frame chứa các trường nhập liệu
        input_frame = tk.Frame(pwd_window)
        input_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Mật khẩu mới
        tk.Label(input_frame, text="Mật khẩu mới\n新密码:", font=("Arial", 12), anchor="w").grid(row=0, column=0,sticky="w", pady=5)
        new_password_entry = tk.Entry(input_frame, font=("Arial", 12), width=25, show="*")
        new_password_entry.grid(row=0, column=1, sticky="w", pady=5)

        # Xác nhận mật khẩu
        tk.Label(input_frame, text="Xác nhận mật khẩu\n确认密码:", font=("Arial", 12), anchor="w").grid(row=1, column=0,sticky="w",pady=5)
        confirm_password_entry = tk.Entry(input_frame, font=("Arial", 12), width=25, show="*")
        confirm_password_entry.grid(row=1, column=1, sticky="w", pady=5)

        # Nút Cập nhật
        update_button = tk.Button(
            input_frame,
            text="CẬP NHẬT\n更新",
            font=("Arial", 12, "bold"),
            bg="#3498db",
            fg="white",
            command=lambda: self.update_user_password(pwd_window, user_id, new_password_entry.get(),confirm_password_entry.get())
        )
        update_button.grid(row=2, column=0, columnspan=2, pady=20)

    def update_user_password(self, window, user_id, new_password, confirm_password):
        """Cập nhật mật khẩu mới cho người dùng"""
        # Kiểm tra thông tin nhập
        if not new_password:
            messagebox.showerror("Lỗi", "Vui lòng nhập mật khẩu mới!\n请输入新密码!")
            return

        if new_password != confirm_password:
            messagebox.showerror("Lỗi", "Mật khẩu xác nhận không khớp!\n确认密码不匹配!")
            return

        # Cập nhật mật khẩu
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            cursor.execute("UPDATE users SET password = ? WHERE id = ?", (new_password, user_id))
            conn.commit()

            # Đóng cửa sổ
            window.destroy()

            messagebox.showinfo("Thành công", "Đã cập nhật mật khẩu thành công!\n密码已成功更新!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể cập nhật mật khẩu: {str(e)}")
        finally:
            conn.close()

    def open_registration_form(self):
        """Mở form đăng ký khuôn mặt"""

        # Kiểm tra quyền admin
        if not self.is_admin:
            messagebox.showwarning("Cảnh báo", "Bạn không có quyền sử dụng chức năng này!\n您无权使用此功能!")
            return

        # Tạo cửa sổ đăng ký
        reg_window = tk.Toplevel(self.root)
        reg_window.title("Đăng Ký Khuôn Mặt人脸注册")
        reg_window.geometry("600x500")
        reg_window.resizable(False, False)

        # Frame tiêu đề
        title_frame = tk.Frame(reg_window, bg="#2E86C1")
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="ĐĂNG KÝ KHUÔN MẶT\n人脸注册",
            font=("Arial", 16, "bold"),
            bg="#2E86C1",
            fg="white",
            pady=10
        ).pack()

        # Frame chứa form
        form_frame = tk.Frame(reg_window, padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Label và Entry cho thông tin
        # Họ và tên
        tk.Label(form_frame, text="Họ và Tên\n姓名:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10,sticky="w")
        name_entry = tk.Entry(form_frame, width=30, font=("Arial", 12))
        name_entry.grid(row=0, column=1, padx=10, pady=10)

        # Mã nhân viên
        tk.Label(form_frame, text="Mã nhân viên\n员工编号:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10,sticky="w")
        employee_id_entry = tk.Entry(form_frame, width=30, font=("Arial", 12))
        employee_id_entry.grid(row=1, column=1, padx=10, pady=10)

        # Bộ phận
        tk.Label(form_frame, text="Bộ phận\n部门:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=10,sticky="w")
        department_entry = tk.Entry(form_frame, width=30, font=("Arial", 12))
        department_entry.grid(row=2, column=1, padx=10, pady=10)

        # Quyền hạn
        tk.Label(form_frame, text="Quyền hạn\n权限:", font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=10,sticky="w")
        permission_combo = ttk.Combobox(form_frame, values=["Level 1", "Level 2", "Level 3"], width=28,
                                        font=("Arial", 12))
        permission_combo.grid(row=3, column=1, padx=10, pady=10)
        permission_combo.current(0)

        # Khu vực làm việc
        tk.Label(form_frame, text="Khu vực làm việc\n工作区域:", font=("Arial", 12)).grid(row=4, column=0, padx=10,pady=10, sticky="w")

        # Frame chứa các checkbox khu vực
        areas_frame = tk.Frame(form_frame)
        areas_frame.grid(row=4, column=1, padx=10, pady=10, sticky="w")

        # Các biến lưu trạng thái checkbox
        area_vars = {
            "A": tk.IntVar(),
            "B": tk.IntVar(),
            "C": tk.IntVar(),
            "D": tk.IntVar(),
            "E": tk.IntVar(),
            "Thí nghiệm": tk.IntVar()
        }

        # Tạo các checkbox
        area_checkboxes = []
        for i, (area, var) in enumerate(area_vars.items()):
            checkbox = tk.Checkbutton(areas_frame, text=area, variable=var, font=("Arial", 11))
            checkbox.grid(row=i // 3, column=i % 3, sticky="w", padx=5)
            area_checkboxes.append(checkbox)

        # Nút Đăng Ký
        register_button = tk.Button(
            form_frame,
            text="ĐĂNG KÝ\n注册",
            font=("Arial", 14, "bold"),
            bg="#27ae60",
            fg="white",
            command=lambda: self.start_face_capture(
                reg_window,
                name_entry.get(),
                employee_id_entry.get(),
                department_entry.get(),
                permission_combo.get(),
                {area: var.get() for area, var in area_vars.items()}
            )
        )
        register_button.grid(row=5, column=0, columnspan=2, pady=20)

    def start_face_capture(self, reg_window, name, employee_id, department, permission, work_areas):
        """Bắt đầu chụp ảnh khuôn mặt với lật khung hình để không bị ngược"""
        # Kiểm tra thông tin nhập
        if not name or not employee_id or not department or not permission:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!\n请填写完整信息!")
            return

        # Kiểm tra ít nhất một khu vực được chọn
        if not any(work_areas.values()):
            messagebox.showerror("Lỗi", "Vui lòng chọn ít nhất một khu vực làm việc!\n请至少选择一个工作区域!")
            return

        # Kiểm tra mã nhân viên đã tồn tại chưa
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM employees WHERE employee_id = ?", (employee_id,))
        if cursor.fetchone()[0] > 0:
            conn.close()
            messagebox.showerror("Lỗi", f"Mã nhân viên '{employee_id}' đã tồn tại!\n员工编号已存在!")
            return

        conn.close()

        # Đóng cửa sổ đăng ký
        reg_window.destroy()

        # Tạo thư mục lưu ảnh cho nhân viên
        folder_name = f"{employee_id}"
        employee_folder = os.path.join(self.employee_images_dir, folder_name)
        os.makedirs(employee_folder, exist_ok=True)

        # Mở cửa sổ camera
        capture_window = tk.Toplevel(self.root)
        capture_window.title("Chụp ảnh khuôn mặt\n拍摄人脸照片")
        capture_window.geometry("700x600")

        # Frame hiển thị camera
        camera_frame = tk.Frame(capture_window)
        camera_frame.pack(pady=10, expand=True, fill=tk.BOTH)

        # Label hiển thị camera
        camera_label = tk.Label(camera_frame)
        camera_label.pack(expand=True, fill=tk.BOTH)

        # Label hiển thị số ảnh đã chụp
        count_label = tk.Label(capture_window, text="Chuẩn bị chụp ảnh...\n准备拍照...", font=("Arial", 14))
        count_label.pack(pady=10)

        # Thiết lập camera
        self.cap = cv2.VideoCapture(self.camera_index)

        if not self.cap.isOpened():
            messagebox.showerror("Lỗi", "Không thể kết nối camera!\n无法连接摄像头!")
            capture_window.destroy()
            return

        # Biến đếm số ảnh đã chụp
        photo_count = 0
        max_photos = 30
        start_time = time.time()
        duration = 7  # Thời gian chụp (giây)

        # Hàm chụp và lưu ảnh
        def capture_images():
            nonlocal photo_count

            ret, frame = self.cap.read()
            if not ret:
                messagebox.showerror("Lỗi", "Không thể đọc dữ liệu từ camera!\n无法从摄像头读取数据!")
                self.cap.release()
                capture_window.destroy()
                return

            # Lật khung hình theo trục ngang để không bị ngược
            frame = cv2.flip(frame, 1)

            # Hiển thị frame
            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(cv2image)
            imgtk = ImageTk.PhotoImage(image=img)
            camera_label.imgtk = imgtk
            camera_label.configure(image=imgtk)

            # Kiểm tra thời gian
            elapsed_time = time.time() - start_time
            remaining_time = max(0, duration - elapsed_time)

            # Nếu còn thời gian và chưa đủ số ảnh
            if elapsed_time <= duration and photo_count < max_photos:
                # Lưu ảnh đã lật
                img_path = os.path.join(employee_folder, f"{photo_count + 1}.jpg")
                cv2.imwrite(img_path, frame)
                photo_count += 1

                # Cập nhật label
                count_label.config(text=f"Đã chụp: {photo_count}/{max_photos} ảnh - Còn lại: {remaining_time:.1f}s")

                # Tiếp tục chụp
                capture_window.after(200, capture_images)
            else:
                # Kết thúc chụp ảnh
                self.cap.release()

                # Tạo vector khuôn mặt
                face_encodings = self.create_face_encodings(employee_folder)

                if face_encodings:
                    # Lưu thông tin nhân viên vào cơ sở dữ liệu
                    self.save_employee_to_database(
                        name, employee_id, department, permission, folder_name,
                        face_encodings, work_areas
                    )

                    messagebox.showinfo("Thành công", "ĐĂNG KÝ KHUÔN MẶT THÀNH CÔNG\n人脸注册成功")
                else:
                    messagebox.showerror("Lỗi", "Không thể nhận diện khuôn mặt trong ảnh chụp!\n无法识别人脸")

                    # Xóa thư mục chứa ảnh đã chụp
                    shutil.rmtree(employee_folder, ignore_errors=True)

                capture_window.destroy()

        # Bắt đầu chụp ảnh
        capture_window.after(500, capture_images)

    def create_face_encodings(self, folder_path):
        """Tạo vector đặc trưng khuôn mặt từ các ảnh"""
        face_encodings = []

        # Lấy danh sách ảnh trong thư mục
        image_files = [f for f in os.listdir(folder_path) if f.endswith('.jpg')]

        # Trích xuất đặc trưng khuôn mặt từ mỗi ảnh
        for image_file in image_files:
            img_path = os.path.join(folder_path, image_file)
            image = face_recognition.load_image_file(img_path)

        # Tìm tất cả khuôn mặt trong ảnh
        face_locations = face_recognition.face_locations(image)

        if face_locations:
            # Lấy encoding của khuôn mặt đầu tiên
            encoding = face_recognition.face_encodings(image, [face_locations[0]])[0]
            # Lưu trực tiếp numpy array, không chuyển sang list
            face_encodings.append(encoding)

        return face_encodings

    def save_employee_to_database(self, name, employee_id, department, permission, image_folder, face_encodings,work_areas):
        """Lưu thông tin nhân viên và vector khuôn mặt vào cơ sở dữ liệu SQLite"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            # Bắt đầu transaction
            conn.execute("BEGIN TRANSACTION")

            # Lưu thông tin nhân viên
            cursor.execute("""
            INSERT INTO employees (name, employee_id, department, permission, image_folder)
            VALUES (?, ?, ?, ?, ?)
            """, (name, employee_id, department, permission, image_folder))

            # Lưu các vector khuôn mặt
            for encoding in face_encodings:
                # Chuyển numpy array thành binary blob
                encoding_blob = encoding.tobytes()

                cursor.execute("""
                INSERT INTO face_encodings (employee_id, encoding)
                VALUES (?, ?)
                """, (employee_id, encoding_blob))

            # Lưu thông tin khu vực làm việc
            for area, is_selected in work_areas.items():
                if is_selected:
                    cursor.execute("""
                    INSERT INTO work_areas (employee_id, area)
                    VALUES (?, ?)
                """, (employee_id, area))

            # Commit transaction
            conn.commit()

        except Exception as e:
            # Rollback nếu có lỗi
            conn.rollback()
            print(f"Database error: {str(e)}")  # In lỗi ra console để debug
            messagebox.showerror("Lỗi", f"Không thể lưu thông tin nhân viên: {str(e)}")
        finally:
            conn.close()

    def open_face_verification(self):
        """Mở cửa sổ xác thực khuôn mặt - tối ưu mượt + phân quyền"""
        if self.is_capturing:
            return

        # Kiểm tra dữ liệu nhân viên
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM employees")
        if cursor.fetchone()[0] == 0:
            conn.close()
            messagebox.showwarning("Cảnh báo", "Chưa có nhân viên nào được đăng ký!")
            return
        conn.close()

        self.is_capturing = True

        verify_window = tk.Toplevel(self.root)
        verify_window.title("Đang xác thực khuôn mặt - Programmed IT Việt Nam PMP - @2025")
        verify_window.geometry("800x700")

        camera_label = tk.Label(verify_window)
        camera_label.pack(expand=True, fill=tk.BOTH)

        # Thêm dòng chữ "Programmed IT Việt Nam PMP - @2025" ở dưới cùng
        footer_label = tk.Label(
            verify_window,
            text="Programmed IT Việt Nam PMP - @2025",
            font=("Arial", 10),
            fg="#666666",
            pady=10
        )
        footer_label.pack(side=tk.BOTTOM)

        self.cap = cv2.VideoCapture(self.camera_index)
        if not self.cap.isOpened():
            messagebox.showerror("Lỗi", f"Không thể mở camera {self.camera_index}!\n无法打开摄像头!")
            self.is_capturing = False
            verify_window.destroy()
            return

        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 480)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 360)

        # Tải dữ liệu khuôn mặt đã mã hóa
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT e.employee_id, e.name, e.department, e.permission, f.encoding
            FROM employees e
            JOIN face_encodings f ON e.employee_id = f.employee_id
        """)
        self.known_faces = []
        for row in cursor.fetchall():
            eid, name, dept, perm, enc_blob = row
            face_enc = np.frombuffer(enc_blob, dtype=np.float64)
            cursor.execute("SELECT area FROM work_areas WHERE employee_id = ?", (eid,))
            areas = [r[0] for r in cursor.fetchall()]
            self.known_faces.append({
                "employee_id": eid,
                "name": name,
                "department": dept,
                "permission": perm,
                "face_encoding": face_enc,
                "areas": areas
            })
        conn.close()

        # Biến xử lý
        self.last_faces = []
        self.frame_count = 0
        self.process_interval = 5
        self.last_detected = {}

        # Luồng 1: Cập nhật video mượt
        def update_camera():
            if not self.is_capturing or not self.cap.isOpened():
                return

            ret, frame = self.cap.read()
            if not ret:
                camera_label.after(30, update_camera)
                return

            # Lật khung hình theo trục ngang để không bị ngược
            frame = cv2.flip(frame, 1)

            display_frame = frame.copy()

            for face in self.last_faces:
                top, right, bottom, left = face["location"]
                name = face["name"]
                color = face["color"]
                cv2.rectangle(display_frame, (left, top), (right, bottom), color, 2)
                cv2.putText(display_frame, name, (left, top - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)

            img = cv2.cvtColor(display_frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(img)
            imgtk = ImageTk.PhotoImage(image=img)
            camera_label.imgtk = imgtk
            camera_label.configure(image=imgtk)

            camera_label.after(30, update_camera)

        # Luồng 2: Nhận diện khuôn mặt
        def recognize_faces():
            if not self.is_capturing or not self.cap.isOpened():
                return

            ret, frame = self.cap.read()
            if not ret:
                camera_label.after(200, recognize_faces)
                return

            # Lật khung hình theo trục ngang để không bị ngược
            frame = cv2.flip(frame, 1)

            # Áp dụng bộ lọc Gaussian để giảm nhiễu
            frame = cv2.GaussianBlur(frame, (5, 5), 0)

            small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
            rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
            face_locations = face_recognition.face_locations(rgb_small_frame, model='hog')

            self.last_faces.clear()

            if face_locations:
                # Xử lý tất cả khuôn mặt trong khung hình
                encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
                for (face_location, face_encoding) in zip(face_locations, encodings):
                    match = None
                    for face in self.known_faces:
                        is_match = face_recognition.compare_faces(
                            [face["face_encoding"]], face_encoding, tolerance=0.40
                        )[0]
                        if is_match:
                            match = face
                            break

                    top, right, bottom, left = [v * 4 for v in face_location]

                    if match:
                        authorized = self.current_work_area in match["areas"]
                        level_colors = {
                            "Level 1": (0, 165, 255),  # Cam
                            "Level 2": (255, 0, 0),  # Xanh dương
                            "Level 3": (0, 255, 0)  # Xanh lá
                        }
                        color = level_colors.get(match["permission"], (255, 255, 255))
                        if not authorized:
                            color = (0, 0, 255)

                        name = f"{match['employee_id']} - {match['permission']}"
                        now = time.time()
                        if match["employee_id"] not in self.last_detected or now - self.last_detected[
                            match["employee_id"]] > 60:
                            self.record_verification_history(match, authorized)
                            self.last_detected[match["employee_id"]] = now
                    else:
                        name = "Not registered"
                        color = (0, 0, 255)

                    self.last_faces.append({
                        "location": (top, right, bottom, left),
                        "name": name,
                        "color": color
                    })

            camera_label.after(200, recognize_faces)

        # Bắt đầu cả hai luồng
        update_camera()
        recognize_faces()

        verify_window.protocol("WM_DELETE_WINDOW", lambda: self.stop_verification(verify_window))

    def stop_verification(self, window):
        """Dừng xác thực và đóng cửa sổ"""
        self.is_capturing = False
        # Đợi thread kết thúc
        time.sleep(0.5)
        window.destroy()

    def record_verification_history(self, employee, is_authorized):
        """Ghi nhận lịch sử xác thực vào cơ sở dữ liệu"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            # Thêm bản ghi lịch sử
            cursor.execute("""
                INSERT INTO history 
                (employee_id, name, department, permission, area, is_authorized) 
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                employee["employee_id"],
                employee["name"],
                employee["department"],
                employee["permission"],
                self.current_work_area,
                1 if is_authorized else 0
            ))

            conn.commit()
        except Exception as e:
            print(f"Lỗi khi ghi lịch sử: {e}")
        finally:
            conn.close()

    def show_employee_list(self):
        """Hiển thị danh sách nhân viên"""
        # Tạo cửa sổ danh sách nhân viên
        list_window = tk.Toplevel(self.root)
        list_window.title("Danh Sách Nhân Viên员工名单")
        list_window.geometry("900x600")

        # Frame tiêu đề
        title_frame = tk.Frame(list_window, bg="#1f618d")
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="DANH SÁCH NHÂN VIÊN\n员工名单",
            font=("Arial", 16, "bold"),
            bg="#1f618d",
            fg="white",
            pady=10
        ).pack()

        # Frame tìm kiếm
        search_frame = tk.Frame(list_window)
        search_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(search_frame, text="Tìm kiếm:搜索", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        search_entry = tk.Entry(search_frame, width=40, font=("Arial", 12))
        search_entry.pack(side=tk.LEFT, padx=5)

        search_button = tk.Button(
            search_frame,
            text="Tìm查找",
            font=("Arial", 12),
            bg="#3498db",
            fg="white",
            command=lambda: self.search_employees(treeview, search_entry.get())
        )
        search_button.pack(side=tk.LEFT, padx=5)

        # Treeview hiển thị danh sách
        columns = ("name", "employee_id", "department", "permission", "work_areas", "created_at")
        treeview = ttk.Treeview(list_window, columns=columns, show="headings")

        # Thiết lập tiêu đề cột
        treeview.heading("name", text="Họ và Tên姓名")
        treeview.heading("employee_id", text="Mã Nhân Viên员工编号")
        treeview.heading("department", text="Bộ Phận部门")
        treeview.heading("permission", text="Quyền Hạn权限")
        treeview.heading("work_areas", text="Khu Vực工作区域")
        treeview.heading("created_at", text="Ngày Tạo注册日期")

        # Thiết lập chiều rộng cột
        treeview.column("name", width=150)
        treeview.column("employee_id", width=100)
        treeview.column("department", width=150)
        treeview.column("permission", width=100)
        treeview.column("work_areas", width=150)
        treeview.column("created_at", width=150)

        # Thêm dữ liệu vào treeview
        self.populate_employee_treeview(treeview)

        # Scrollbar
        scrollbar = ttk.Scrollbar(list_window, orient=tk.VERTICAL, command=treeview.yview)
        treeview.configure(yscrollcommand=scrollbar.set)

        # Đặt treeview và scrollbar
        treeview.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Frame chứa các nút
        button_frame = tk.Frame(list_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        # Nút Xóa nhân viên
        delete_button = tk.Button(
            button_frame,
            text="Xóa\n删除",
            font=("Arial", 12),
            bg="#e74c3c",
            fg="white",
            command=lambda: self.delete_employee(treeview)
        )
        delete_button.pack(side=tk.LEFT, padx=5)

        # Nút Sửa thông tin
        edit_button = tk.Button(
            button_frame,
            text="Sửa\n修改",
            font=("Arial", 12),
            bg="#f39c12",
            fg="white",
            command=lambda: self.edit_employee(treeview)
        )
        edit_button.pack(side=tk.LEFT, padx=5)

    def populate_employee_treeview(self, treeview):
        """Cập nhật danh sách nhân viên vào treeview"""
        # Xóa dữ liệu cũ
        for item in treeview.get_children():
            treeview.delete(item)

        # Lấy dữ liệu từ SQLite
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Query để lấy thông tin nhân viên và khu vực làm việc
        # Thay đổi thứ tự của created_at và work_areas trong câu truy vấn
        cursor.execute("""
            SELECT e.name, e.employee_id, e.department, e.permission, 
                GROUP_CONCAT(w.area, ', ') as work_areas, e.created_at
            FROM employees e
            LEFT JOIN work_areas w ON e.employee_id = w.employee_id
            GROUP BY e.employee_id
            ORDER BY e.name
        """)

        # Thêm dữ liệu mới
        for row in cursor.fetchall():
            treeview.insert("", tk.END, values=row)

        conn.close()

    def search_employees(self, treeview, keyword):
        """Tìm kiếm nhân viên"""
        # Xóa dữ liệu cũ
        for item in treeview.get_children():
            treeview.delete(item)

        # Nếu không có từ khóa, hiển thị tất cả
        if not keyword:
            self.populate_employee_treeview(treeview)
            return

        # Tìm kiếm theo từ khóa
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Query tìm kiếm - Đã sửa thứ tự để khớp với populate_employee_treeview
        cursor.execute("""
            SELECT e.name, e.employee_id, e.department, e.permission, 
                GROUP_CONCAT(w.area, ', ') as work_areas, e.created_at
            FROM employees e
            LEFT JOIN work_areas w ON e.employee_id = w.employee_id
            WHERE e.name LIKE ? OR e.employee_id LIKE ? OR e.department LIKE ? OR e.permission LIKE ?
            GROUP BY e.employee_id
            ORDER BY e.name
        """, (f"%{keyword}%", f"%{keyword}%", f"%{keyword}%", f"%{keyword}%"))

        # Thêm kết quả tìm kiếm
        for row in cursor.fetchall():
            treeview.insert("", tk.END, values=row)

        conn.close()

    def delete_employee(self, treeview):
        """Xóa nhân viên"""

        # Kiểm tra quyền admin
        if not self.is_admin:
            messagebox.showwarning("Cảnh báo", "Bạn không có quyền xóa thông tin nhân viên!\n您无权删除员工信息!")
            return

        # Lấy item được chọn
        selected_item = treeview.selection()

        if not selected_item:
            messagebox.showwarning("Cảnh báo 警告", "Vui lòng chọn nhân viên cần xóa!\n选择要删除的员工!")
            return

        # Lấy item được chọn
        selected_item = treeview.selection()

        if not selected_item:
            messagebox.showwarning("Cảnh báo 警告", "Vui lòng chọn nhân viên cần xóa!\n选择要删除的员工!")
            return

        # Lấy thông tin nhân viên
        values = treeview.item(selected_item[0], "values")
        employee_id = values[1]

        # Xác nhận xóa
        confirm = messagebox.askyesno("Xác nhận 重新询问",f"Bạn có chắc muốn xóa nhân viên 确定要删除该员工吗？ {values[0]}?")

        if confirm:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            try:
                # Bắt đầu transaction
                conn.execute("BEGIN TRANSACTION")

                # Lấy tên thư mục ảnh
                cursor.execute("SELECT image_folder FROM employees WHERE employee_id = ?", (employee_id,))
                folder_name = cursor.fetchone()[0]

                # Xóa dữ liệu từ các bảng
                cursor.execute("DELETE FROM work_areas WHERE employee_id = ?", (employee_id,))
                cursor.execute("DELETE FROM face_encodings WHERE employee_id = ?", (employee_id,))
                cursor.execute("DELETE FROM employees WHERE employee_id = ?", (employee_id,))

                # Commit transaction
                conn.commit()

                # Xóa thư mục ảnh
                folder_path = os.path.join(self.employee_images_dir, folder_name)
                if os.path.exists(folder_path):
                    shutil.rmtree(folder_path, ignore_errors=True)

                # Cập nhật treeview
                treeview.delete(selected_item[0])

                messagebox.showinfo("Thành công 完成", "Đã xóa nhân viên thành công!\n已成功删除员工")

            except Exception as e:
                # Rollback nếu có lỗi
                conn.rollback()
                messagebox.showerror("Lỗi", f"Không thể xóa nhân viên: {str(e)}")
            finally:
                conn.close()

    def edit_employee(self, treeview):
        """Sửa thông tin nhân viên"""

        # Kiểm tra quyền admin
        if not self.is_admin:
            messagebox.showwarning("Cảnh báo", "Bạn không có quyền sửa thông tin nhân viên!\n您无权修改员工信息!")
            return

        # Lấy item được chọn
        selected_item = treeview.selection()

        if not selected_item:
            messagebox.showwarning("Cảnh báo 警告", "Vui lòng chọn nhân viên cần sửa!\n请选择需要修改的员工")
            return

        if not selected_item:
            messagebox.showwarning("Cảnh báo 警告", "Vui lòng chọn nhân viên cần sửa!\n请选择需要修改的员工")
            return
        # Lấy thông tin nhân viên
        values = treeview.item(selected_item[0], "values")
        name = values[0]
        employee_id = values[1]
        department = values[2]
        permission = values[3]

        # Lấy danh sách khu vực làm việc
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT area FROM work_areas WHERE employee_id = ?", (employee_id,))
        work_areas = [row[0] for row in cursor.fetchall()]
        conn.close()

        # Tạo cửa sổ chỉnh sửa
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Sửa Thông Tin Nhân Viên 修改员工信息")
        edit_window.geometry("600x500")

        # Frame tiêu đề
        title_frame = tk.Frame(edit_window, bg="#f39c12")
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="SỬA THÔNG TIN NHÂN VIÊN\n修改员工信息",
            font=("Arial", 16, "bold"),
            bg="#f39c12",
            fg="white",
            pady=10
        ).pack()

        # Frame chứa form
        form_frame = tk.Frame(edit_window, padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Label và Entry cho thông tin
        # Họ và tên
        tk.Label(form_frame, text="Họ và Tên\n姓名:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10,sticky="w")
        name_entry = tk.Entry(form_frame, width=30, font=("Arial", 12))
        name_entry.insert(0, name)
        name_entry.grid(row=0, column=1, padx=10, pady=10)

        # Mã nhân viên (không cho phép sửa)
        tk.Label(form_frame, text="Mã nhân viên\n员工编号:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10,sticky="w")
        employee_id_entry = tk.Entry(form_frame, width=30, font=("Arial", 12))
        employee_id_entry.insert(0, employee_id)
        employee_id_entry.config(state="readonly")
        employee_id_entry.grid(row=1, column=1, padx=10, pady=10)

        # Bộ phận
        tk.Label(form_frame, text="Bộ phận\n部门:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=10,sticky="w")
        department_entry = tk.Entry(form_frame, width=30, font=("Arial", 12))
        department_entry.insert(0, department)
        department_entry.grid(row=2, column=1, padx=10, pady=10)

        # Quyền hạn
        tk.Label(form_frame, text="Quyền hạn\n权限:", font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=10,sticky="w")
        permission_combo = ttk.Combobox(form_frame, values=["Level 1", "Level 2", "Level 3"], width=28,
                                        font=("Arial", 12))
        permission_combo.set(permission)
        permission_combo.grid(row=3, column=1, padx=10, pady=10)

        # Khu vực làm việc
        tk.Label(form_frame, text="Khu vực làm việc\n工作区域:", font=("Arial", 12)).grid(row=4, column=0, padx=10,pady=10, sticky="w")

        # Frame chứa các checkbox khu vực
        areas_frame = tk.Frame(form_frame)
        areas_frame.grid(row=4, column=1, padx=10, pady=10, sticky="w")

        # Các biến lưu trạng thái checkbox
        area_vars = {
            "A": tk.IntVar(value=1 if "A" in work_areas else 0),
            "B": tk.IntVar(value=1 if "B" in work_areas else 0),
            "C": tk.IntVar(value=1 if "C" in work_areas else 0),
            "D": tk.IntVar(value=1 if "D" in work_areas else 0),
            "E": tk.IntVar(value=1 if "E" in work_areas else 0),
            "Thí nghiệm": tk.IntVar(value=1 if "Thí nghiệm" in work_areas else 0)
        }

        # Tạo các checkbox
        area_checkboxes = []
        for i, (area, var) in enumerate(area_vars.items()):
            checkbox = tk.Checkbutton(areas_frame, text=area, variable=var, font=("Arial", 11))
            checkbox.grid(row=i // 3, column=i % 3, sticky="w", padx=5)
            area_checkboxes.append(checkbox)

        # Nút Cập Nhật
        update_button = tk.Button(
            form_frame,
            text="CẬP NHẬT\n更新",
            font=("Arial", 14, "bold"),
            bg="#27ae60",
            fg="white",
            command=lambda: self.update_employee_data(
                edit_window,
                treeview,
                selected_item[0],
                employee_id,
                name_entry.get(),
                department_entry.get(),
                permission_combo.get(),
                {area: var.get() for area, var in area_vars.items()}
            )
        )
        update_button.grid(row=5, column=0, columnspan=2, pady=20)

    def update_employee_data(self, window, treeview, item, employee_id, name, department, permission, work_areas):
        """Cập nhật thông tin nhân viên"""
        # Kiểm tra thông tin nhập
        if not name or not department or not permission:
            messagebox.showerror("Lỗi\n错误", "Vui lòng nhập đầy đủ thông tin!\n填写完整信息!")
            return

        # Kiểm tra ít nhất một khu vực được chọn
        if not any(work_areas.values()):
            messagebox.showerror("Lỗi", "Vui lòng chọn ít nhất một khu vực làm việc!\n请至少选择一个工作区域!")
            return

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            # Bắt đầu transaction
            conn.execute("BEGIN TRANSACTION")

            # Cập nhật thông tin nhân viên
            cursor.execute("""
                UPDATE employees 
                SET name = ?, department = ?, permission = ?
                WHERE employee_id = ?
            """, (name, department, permission, employee_id))

            # Xóa tất cả khu vực hiện tại
            cursor.execute("DELETE FROM work_areas WHERE employee_id = ?", (employee_id,))

            # Thêm lại các khu vực mới
            for area, is_selected in work_areas.items():
                if is_selected:
                    cursor.execute("""
                        INSERT INTO work_areas (employee_id, area)
                        VALUES (?, ?)
                    """, (employee_id, area))

            # Commit transaction
            conn.commit()

            # Đóng cửa sổ
            window.destroy()

            # Cập nhật lại treeview
            self.populate_employee_treeview(treeview)

            messagebox.showinfo("Thành công 完成", "Đã cập nhật thông tin nhân viên!\n已更新员工信息!")

        except Exception as e:
            # Rollback nếu có lỗi
            conn.rollback()
            messagebox.showerror("Lỗi", f"Không thể cập nhật thông tin nhân viên: {str(e)}")
        finally:
            conn.close()

    def show_history(self):
        """Hiển thị lịch sử xác thực"""
        # Tạo cửa sổ lịch sử
        history_window = tk.Toplevel(self.root)
        history_window.title("Lịch Sử Xác Thực 验证历史")
        history_window.geometry("900x600")

        # Frame tiêu đề
        title_frame = tk.Frame(history_window, bg="#2874a6")
        title_frame.pack(fill=tk.X)

        tk.Label(
            title_frame,
            text="LỊCH SỬ XÁC THỰC\n验证历史",
            font=("Arial", 16, "bold"),
            bg="#2874a6",
            fg="white",
            pady=10
        ).pack()

        # Frame tìm kiếm và bộ lọc
        filter_frame = tk.Frame(history_window)
        filter_frame.pack(fill=tk.X, padx=10, pady=10)

        # Các điều kiện lọc
        tk.Label(filter_frame, text="Tìm kiếm:\n搜索:", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5)
        search_entry = tk.Entry(filter_frame, width=25, font=("Arial", 12))
        search_entry.grid(row=0, column=1, padx=5, pady=5)

        # Bộ lọc ngày tháng
        tk.Label(filter_frame, text="Từ ngày:\n从日期:", font=("Arial", 12)).grid(row=0, column=2, padx=5, pady=5)
        from_date_entry = tk.Entry(filter_frame, width=15, font=("Arial", 12))
        from_date_entry.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(filter_frame, text="Đến ngày:\n至日期:", font=("Arial", 12)).grid(row=0, column=4, padx=5, pady=5)
        to_date_entry = tk.Entry(filter_frame, width=15, font=("Arial", 12))
        to_date_entry.grid(row=0, column=5, padx=5, pady=5)

        # Bộ lọc khu vực
        tk.Label(filter_frame, text="Khu vực:\n区域:", font=("Arial", 12)).grid(row=1, column=0, padx=5, pady=5)
        area_combo = ttk.Combobox(filter_frame, values=["Tất cả", "A", "B", "C", "D", "E", "Thí nghiệm"], width=12,
                                  font=("Arial", 12))
        area_combo.grid(row=1, column=1, padx=5, pady=5)
        area_combo.current(0)

        # Bộ lọc trạng thái
        tk.Label(filter_frame, text="Trạng thái:\n状态:", font=("Arial", 12)).grid(row=1, column=2, padx=5, pady=5)
        status_combo = ttk.Combobox(filter_frame, values=["Tất cả", "Có quyền", "Không có quyền"], width=12,
                                    font=("Arial", 12))
        status_combo.grid(row=1, column=3, padx=5, pady=5)
        status_combo.current(0)

        # Nút tìm kiếm
        search_button = tk.Button(
            filter_frame,
            text="Tìm kiếm\n搜索",
            font=("Arial", 12),
            bg="#3498db",
            fg="white",
            command=lambda: self.search_history(
                treeview,
                search_entry.get(),
                from_date_entry.get(),
                to_date_entry.get(),
                area_combo.get(),
                status_combo.get()
            )
        )
        search_button.grid(row=1, column=5, padx=5, pady=5, sticky="e")

        # Treeview hiển thị lịch sử
        columns = ("id", "name", "employee_id", "department", "permission", "area", "is_authorized", "timestamp")
        treeview = ttk.Treeview(history_window, columns=columns, show="headings")

        # Thiết lập tiêu đề cột
        treeview.heading("id", text="ID")
        treeview.heading("name", text="Họ và Tên\n姓名")
        treeview.heading("employee_id", text="Mã NV\n工号")
        treeview.heading("department", text="Bộ Phận\n部门")
        treeview.heading("permission", text="Quyền Hạn\n权限")
        treeview.heading("area", text="Khu Vực\n区域")
        treeview.heading("is_authorized", text="Trạng Thái\n状态")
        treeview.heading("timestamp", text="Thời Gian\n时间")

        # Thiết lập chiều rộng cột
        treeview.column("id", width=50)
        treeview.column("name", width=150)
        treeview.column("employee_id", width=80)
        treeview.column("department", width=100)
        treeview.column("permission", width=80)
        treeview.column("area", width=80)
        treeview.column("is_authorized", width=100)
        treeview.column("timestamp", width=150)

        # Thêm dữ liệu vào treeview
        self.populate_history_treeview(treeview)

        # Scrollbar
        scrollbar = ttk.Scrollbar(history_window, orient=tk.VERTICAL, command=treeview.yview)
        treeview.configure(yscrollcommand=scrollbar.set)

        # Đặt treeview và scrollbar
        treeview.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Nút xuất Excel
        export_button = tk.Button(
            history_window,
            text="Xuất Excel\n导出Excel",
            font=("Arial", 12),
            bg="#16a085",
            fg="white",
            command=lambda: self.export_history_excel()
        )
        export_button.pack(side=tk.RIGHT, padx=10, pady=10)

    def populate_history_treeview(self, treeview):
        """Cập nhật lịch sử vào treeview"""
        # Xóa dữ liệu cũ
        for item in treeview.get_children():
            treeview.delete(item)

        # Lấy dữ liệu từ SQLite
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Query lấy lịch sử xác thực
        cursor.execute("""
            SELECT id, name, employee_id, department, permission, area, is_authorized, timestamp
            FROM history
            ORDER BY timestamp DESC
            LIMIT 1000
        """)

        # Thêm dữ liệu mới
        for row in cursor.fetchall():
            id, name, employee_id, department, permission, area, is_authorized, timestamp = row
            status = "Có quyền" if is_authorized else "Không có quyền"

            treeview.insert("", tk.END, values=(id, name, employee_id, department, permission, area, status, timestamp))

        conn.close()

    def search_history(self, treeview, keyword, from_date, to_date, area, status):
        """Tìm kiếm lịch sử theo các điều kiện"""
        # Xóa dữ liệu cũ
        for item in treeview.get_children():
            treeview.delete(item)

        # Xây dựng câu query
        query = """
            SELECT id, name, employee_id, department, permission, area, is_authorized, timestamp
            FROM history
            WHERE 1=1
        """

        params = []

        # Thêm điều kiện tìm kiếm từ khóa
        if keyword:
            query += """ AND (name LIKE ? OR employee_id LIKE ? OR department LIKE ?)"""
            keyword_param = f"%{keyword}%"
            params.extend([keyword_param, keyword_param, keyword_param])

        # Thêm điều kiện ngày bắt đầu
        if from_date:
            try:
                # Chuyển đổi định dạng ngày tháng
                datetime.datetime.strptime(from_date, "%Y-%m-%d")
                query += """ AND timestamp >= ?"""
                params.append(f"{from_date} 00:00:00")
            except ValueError:
                messagebox.showwarning("Cảnh báo", "Định dạng ngày không hợp lệ (YYYY-MM-DD)!\n日期格式无效!")
                return

        # Thêm điều kiện ngày kết thúc
        if to_date:
            try:
                # Chuyển đổi định dạng ngày tháng
                datetime.datetime.strptime(to_date, "%Y-%m-%d")
                query += """ AND timestamp <= ?"""
                params.append(f"{to_date} 23:59:59")
            except ValueError:
                messagebox.showwarning("Cảnh báo", "Định dạng ngày không hợp lệ (YYYY-MM-DD)!\n日期格式无效!")
                return

        # Thêm điều kiện khu vực
        if area and area != "Tất cả":
            query += """ AND area = ?"""
            params.append(area)

        # Thêm điều kiện trạng thái
        if status and status != "Tất cả":
            if status == "Có quyền":
                query += """ AND is_authorized = 1"""
            elif status == "Không có quyền":
                query += """ AND is_authorized = 0"""

        # Thêm sắp xếp
        query += """ ORDER BY timestamp DESC LIMIT 1000"""

        # Thực hiện truy vấn
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute(query, params)

        # Hiển thị kết quả
        for row in cursor.fetchall():
            id, name, employee_id, department, permission, area, is_authorized, timestamp = row
            status_text = "Có quyền" if is_authorized else "Không có quyền"

            treeview.insert("", tk.END,values=(id, name, employee_id, department, permission, area, status_text, timestamp))

        conn.close()

    def export_history_excel(self):
        """Xuất lịch sử xác thực ra file Excel"""
        # Chọn vị trí lưu file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu Lịch Sử Xác Thực\n保存验证历史"
        )

        if not file_path:
            return

        try:
            # Lấy dữ liệu từ SQLite
            conn = sqlite3.connect(self.db_path)
            query = """
                SELECT name, employee_id, department, permission, area, 
                       CASE WHEN is_authorized = 1 THEN 'Có quyền' ELSE 'Không có quyền' END as status,
                       timestamp
                FROM history
                ORDER BY timestamp DESC
            """

            # Đọc dữ liệu vào DataFrame
            df = pd.read_sql_query(query, conn)

            # Đổi tên cột
            df.columns = [
                "Họ và Tên", "Mã Nhân Viên", "Bộ Phận", "Quyền Hạn",
                "Khu Vực", "Trạng Thái", "Thời Gian"
            ]

            # Xuất ra file Excel
            df.to_excel(file_path, index=False)

            conn.close()

            messagebox.showinfo("Thành công", f"Đã xuất lịch sử xác thực ra file Excel:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất Excel: {str(e)}")

    def export_options(self):
        """Hiển thị các tùy chọn xuất Excel"""
        # Tạo cửa sổ tùy chọn xuất Excel
        export_window = tk.Toplevel(self.root)
        export_window.title("Xuất Excel导出Excel")
        export_window.geometry("400x200")

        # Tiêu đề
        tk.Label(
            export_window,
            text="CHỌN LOẠI DỮ LIỆU CẦN XUẤT\n选择需要导出的数据类型",
            font=("Arial", 14, "bold"),
            pady=10
        ).pack()

        # Nút xuất danh sách nhân viên
        employees_button = tk.Button(
            export_window,
            text="Xuất Danh Sách Nhân Viên\n导出员工名单",
            font=("Arial", 12),
            bg="#3498db",
            fg="white",
            command=lambda: self.export_employees_excel(export_window)
        )
        employees_button.pack(fill=tk.X, padx=20, pady=10)

        # Nút xuất lịch sử
        history_button = tk.Button(
            export_window,
            text="Xuất Lịch Sử Xác Thực\n导出验证历史",
            font=("Arial", 12),
            bg="#2ecc71",
            fg="white",
            command=lambda: self.export_history_excel_from_window(export_window)
        )
        history_button.pack(fill=tk.X, padx=20, pady=10)

    def export_employees_excel(self, parent_window=None):
        """Xuất danh sách nhân viên ra file Excel"""
        # Chọn vị trí lưu file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Lưu Danh Sách Nhân Viên\n保存员工名单"
        )

        if not file_path:
            return

        try:
            # Lấy dữ liệu từ SQLite
            conn = sqlite3.connect(self.db_path)
            # Đã sửa thứ tự của work_areas và created_at
            query = """
                SELECT e.name, e.employee_id, e.department, e.permission, 
                    GROUP_CONCAT(w.area, ', ') as work_areas,
                    e.created_at
                FROM employees e
                LEFT JOIN work_areas w ON e.employee_id = w.employee_id
                GROUP BY e.employee_id
                ORDER BY e.name
            """

            # Đọc dữ liệu vào DataFrame
            df = pd.read_sql_query(query, conn)

            # Đổi tên cột
            df.columns = [
                "Họ và Tên", "Mã Nhân Viên", "Bộ Phận", "Quyền Hạn",
                "Khu Vực Làm Việc", "Ngày Tạo"
            ]

            # Xuất ra file Excel
            df.to_excel(file_path, index=False)

            conn.close()

            # Đóng cửa sổ nếu được mở từ menu xuất Excel
            if parent_window:
                parent_window.destroy()

            messagebox.showinfo("Thành công", f"Đã xuất danh sách nhân viên ra file Excel:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất Excel: {str(e)}")

    def export_history_excel_from_window(self, parent_window):
        """Xuất lịch sử từ cửa sổ menu xuất Excel"""
        parent_window.destroy()
        self.export_history_excel()

        # Hàm main để chạy ứng dụng


def main():
    # Tạo cửa sổ chính
    root = tk.Tk()
    app = FaceRecognitionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
