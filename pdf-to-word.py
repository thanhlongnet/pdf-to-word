import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2docx import Converter
import os
import threading
from tkinter import font as tkfont

class PDFtoWordConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Công Cụ Chuyển PDF Sang Word")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # Thiết lập phông chữ
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.default_font.configure(size=10)
        
        # Màu sắc
        self.bg_color = "#f0f0f0"
        self.button_color = "#4CAF50"
        self.text_color = "#333333"
        
        # Biến lưu đường dẫn file
        self.pdf_path = tk.StringVar()
        self.word_path = tk.StringVar()
        
        # Biến kiểm soát tiến trình
        self.conversion_in_progress = False
        self.progress_window = None
        
        # Tạo giao diện
        self.create_widgets()
    
    def create_widgets(self):
        # Frame chính với màu nền
        main_frame = tk.Frame(
            self.root, 
            padx=30, 
            pady=30,
            bg=self.bg_color
        )
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Tiêu đề
        title_label = tk.Label(
            main_frame, 
            text="CHUYỂN ĐỔI PDF SANG WORD", 
            font=("Arial", 18, "bold"),
            fg="#2c3e50",
            bg=self.bg_color
        )
        title_label.pack(pady=(0, 20))
        
        # Khung chứa nội dung chính
        content_frame = tk.Frame(
            main_frame,
            bg=self.bg_color
        )
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Chọn file PDF
        pdf_frame = tk.Frame(content_frame, bg=self.bg_color)
        pdf_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            pdf_frame, 
            text="File PDF nguồn:", 
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side=tk.LEFT, anchor='w')
        
        pdf_entry = tk.Entry(
            pdf_frame, 
            textvariable=self.pdf_path, 
            width=40,
            font=("Arial", 10),
            bd=2,
            relief=tk.GROOVE
        )
        pdf_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        browse_pdf_btn = tk.Button(
            pdf_frame, 
            text="Chọn file", 
            command=self.browse_pdf,
            bg="#3498db",
            fg="white",
            relief=tk.RAISED,
            bd=2,
            font=("Arial", 10, "bold")
        )
        browse_pdf_btn.pack(side=tk.LEFT)
        
        # Chọn nơi lưu file Word
        word_frame = tk.Frame(content_frame, bg=self.bg_color)
        word_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            word_frame, 
            text="Lưu file Word tại:", 
            font=("Arial", 10),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side=tk.LEFT, anchor='w')
        
        word_entry = tk.Entry(
            word_frame, 
            textvariable=self.word_path, 
            width=40,
            font=("Arial", 10),
            bd=2,
            relief=tk.GROOVE
        )
        word_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        browse_word_btn = tk.Button(
            word_frame, 
            text="Chọn vị trí", 
            command=self.browse_word,
            bg="#3498db",
            fg="white",
            relief=tk.RAISED,
            bd=2,
            font=("Arial", 10, "bold")
        )
        browse_word_btn.pack(side=tk.LEFT)
        
        # Nút chuyển đổi
        convert_btn_frame = tk.Frame(content_frame, bg=self.bg_color)
        convert_btn_frame.pack(pady=30)
        
        self.convert_btn = tk.Button(
            convert_btn_frame, 
            text="CHUYỂN ĐỔI PDF SANG WORD", 
            command=self.start_conversion_thread,
            bg=self.button_color,
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 12, "bold"),
            relief=tk.RAISED,
            bd=3
        )
        self.convert_btn.pack()
        
        # Footer
        footer_frame = tk.Frame(content_frame, bg=self.bg_color)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))
        
        tk.Label(
            footer_frame, 
            text="Công cụ chuyển đổi PDF sang Word - Phiên bản 1.1", 
            font=("Arial", 8),
            bg=self.bg_color,
            fg="#7f8c8d"
        ).pack(side=tk.LEFT)
    
    def browse_pdf(self):
        if self.conversion_in_progress:
            messagebox.showwarning("Thông báo", "Vui lòng chờ quá trình chuyển đổi hiện tại hoàn thành!")
            return
        
        filepath = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filepath:
            self.pdf_path.set(filepath)
            # Tự động đề xuất tên file output
            dirname = os.path.dirname(filepath)
            filename = os.path.splitext(os.path.basename(filepath))[0] + ".docx"
            self.word_path.set(os.path.join(dirname, filename))
    
    def browse_word(self):
        if self.conversion_in_progress:
            messagebox.showwarning("Thông báo", "Vui lòng chờ quá trình chuyển đổi hiện tại hoàn thành!")
            return
        
        filepath = filedialog.asksaveasfilename(
            title="Chọn vị trí lưu file Word",
            defaultextension=".docx",
            filetypes=[("Word Files", "*.docx"), ("All Files", "*.*")]
        )
        if filepath:
            self.word_path.set(filepath)
    
    def start_conversion_thread(self):
        if self.conversion_in_progress:
            return
            
        pdf_file = self.pdf_path.get()
        word_file = self.word_path.get()
        
        if not pdf_file:
            messagebox.showerror("Lỗi", "Vui lòng chọn file PDF trước!")
            return
        
        if not word_file:
            messagebox.showerror("Lỗi", "Vui lòng chọn vị trí lưu file Word!")
            return
        
        # Vô hiệu hóa nút chuyển đổi trong khi đang xử lý
        self.convert_btn.config(state=tk.DISABLED)
        self.conversion_in_progress = True
        
        # Hiển thị cửa sổ tiến trình
        self.show_progress_window()
        
        # Tạo và chạy luồng chuyển đổi
        conversion_thread = threading.Thread(
            target=self.convert_pdf_to_word,
            args=(pdf_file, word_file),
            daemon=True
        )
        conversion_thread.start()
        
        # Kiểm tra tiến trình
        self.root.after(100, self.check_conversion_status, conversion_thread)
    
    def show_progress_window(self):
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Đang xử lý...")
        self.progress_window.geometry("400x150")
        self.progress_window.resizable(False, False)
        
        # Hiển thị tên file đang xử lý (chỉ hiển thị tên file, không hiển thị toàn bộ đường dẫn)
        filename = os.path.basename(self.pdf_path.get())
        tk.Label(
            self.progress_window, 
            text=f"Đang chuyển đổi file: {filename}", 
            font=("Arial", 11)
        ).pack(pady=(10, 5))
        
        tk.Label(
            self.progress_window, 
            text="Vui lòng chờ...", 
            font=("Arial", 10)
        ).pack()
        
        self.progress_bar = ttk.Progressbar(
            self.progress_window, 
            orient=tk.HORIZONTAL, 
            length=300, 
            mode='indeterminate'
        )
        self.progress_bar.pack(pady=20)
        self.progress_bar.start()
        
        # Hiển thị kích thước file
        file_size = os.path.getsize(self.pdf_path.get()) / (1024 * 1024)  # Convert to MB
        size_label = tk.Label(
            self.progress_window,
            text=f"Kích thước file: {file_size:.2f} MB",
            font=("Arial", 9)
        )
        size_label.pack()
    
    def check_conversion_status(self, thread):
        if thread.is_alive():
            # Nếu luồng vẫn đang chạy, kiểm tra lại sau 100ms
            self.root.after(100, self.check_conversion_status, thread)
        else:
            # Khi luồng hoàn thành
            self.conversion_complete()
    
    def convert_pdf_to_word(self, pdf_file, word_file):
        try:
            cv = Converter(pdf_file)
            cv.convert(word_file, start=0, end=None)
            cv.close()
            self.conversion_success = True
        except Exception as e:
            self.conversion_success = False
            self.error_message = str(e)
    
    def conversion_complete(self):
        # Dừng thanh tiến trình
        if self.progress_window:
            self.progress_bar.stop()
            self.progress_window.destroy()
        
        # Kích hoạt lại nút chuyển đổi
        self.convert_btn.config(state=tk.NORMAL)
        self.conversion_in_progress = False
        
        # Hiển thị kết quả
        if hasattr(self, 'conversion_success') and self.conversion_success:
            messagebox.showinfo("Thành công", "Chuyển đổi file thành công!")
        else:
            error_msg = getattr(self, 'error_message', "Lỗi không xác định")
            messagebox.showerror("Lỗi", f"Chuyển đổi thất bại: {error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoWordConverter(root)
    
    # Đặt icon cho ứng dụng (nếu có file icon.ico)
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    root.mainloop()
