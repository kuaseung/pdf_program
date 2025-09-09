import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import shutil
import sys
import os

# --- PDF 압축 핵심 기능 ---
def compress_pdf(input_path, output_path, quality_level):
    """
    Ghostscript를 사용하여 PDF를 압축합니다.
    - quality_level: 'screen', 'ebook', 'printer', 'prepress' 중 하나
    """
    gs_command = find_ghostscript_executable()
    if not gs_command:
        messagebox.showerror(
            "오류", 
            "Ghostscript를 찾을 수 없습니다.\n"
            "프로그램을 사용하려면 Ghostscript를 설치하고\n"
            "PATH 환경 변수에 추가해야 합니다."
        )
        return False

    # Ghostscript 명령어 구성
    command = [
        gs_command,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{quality_level}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={output_path}",
        input_path
    ]

    try:
        # 서브프로세스로 Ghostscript 실행
        # Windows에서는 CREATE_NO_WINDOW 플래그로 콘솔 창 숨김
        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        process = subprocess.run(
            command, 
            check=True, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE,
            startupinfo=startupinfo
        )
        return True
    except FileNotFoundError:
        messagebox.showerror("오류", "Ghostscript를 실행할 수 없습니다. 설치를 확인해주세요.")
        return False
    except subprocess.CalledProcessError as e:
        error_message = f"PDF 압축 중 오류가 발생했습니다.\n{e.stderr.decode('utf-8', errors='ignore')}"
        messagebox.showerror("압축 오류", error_message)
        return False
    except Exception as e:
        messagebox.showerror("알 수 없는 오류", f"예상치 못한 오류가 발생했습니다: {e}")
        return False

def find_ghostscript_executable():
    """
    시스템에서 Ghostscript 실행 파일을 찾습니다.
    (Windows: gswin64c.exe, gswin32c.exe / Linux/macOS: gs)
    """
    if sys.platform == "win32":
        for cmd in ["gswin64c", "gswin32c", "gs"]:
            if shutil.which(cmd):
                return shutil.which(cmd)
    else: # Linux, macOS
        if shutil.which("gs"):
            return shutil.which("gs")
    return None

# --- GUI 애플리케이션 ---
class PDFCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 압축기")
        self.root.geometry("500x420")
        self.root.resizable(True, True)

        self.input_file_path = ""

        # 스타일 설정
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", font=('Helvetica', 10))
        self.style.configure("TLabel", padding=5, font=('Helvetica', 10))
        self.style.configure("TFrame", padding=10)
        self.style.configure("Header.TLabel", font=('Helvetica', 14, 'bold'))

        # 메인 프레임
        main_frame = ttk.Frame(root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- UI 요소 생성 ---
        
        # 1. 제목
        header_label = ttk.Label(main_frame, text="PDF 압축 프로그램", style="Header.TLabel")
        header_label.pack(pady=(0, 20))

        # 2. 파일 선택
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="선택된 파일이 없습니다.", width=40, relief="sunken", anchor="w")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)

        browse_button = ttk.Button(file_frame, text="파일 찾기", command=self.browse_file)
        browse_button.pack(side=tk.RIGHT, padx=(10, 0))

        # 3. 압축 품질 선택
        quality_frame = ttk.LabelFrame(main_frame, text="압축 품질 설정", padding="10 10")
        quality_frame.pack(fill=tk.X, pady=10)

        self.quality_var = tk.StringVar(value="ebook")
        qualities = {
            "낮음 (화면용)": "screen",
            "중간 (전자책)": "ebook",
            "높음 (인쇄용)": "printer",
            "최고 (출판용)": "prepress"
        }

        for text, value in qualities.items():
            rb = ttk.Radiobutton(quality_frame, text=text, variable=self.quality_var, value=value)
            rb.pack(anchor="w", pady=2)

        # 4. 압축 실행 버튼
        compress_button = ttk.Button(main_frame, text="압축 시작", command=self.start_compression)
        compress_button.pack(pady=20, fill=tk.X, ipady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=(("PDF 파일", "*.pdf"), ("모든 파일", "*.*"))
        )
        if file_path:
            self.input_file_path = file_path
            # 파일 이름이 너무 길면 잘라서 표시
            display_name = os.path.basename(file_path)
            if len(display_name) > 35:
                display_name = "..." + display_name[-32:]
            self.file_label.config(text=f" {display_name}")

    def start_compression(self):
        if not self.input_file_path:
            messagebox.showwarning("알림", "먼저 압축할 PDF 파일을 선택해주세요.")
            return

        output_path = filedialog.asksaveasfilename(
            title="압축된 PDF 파일 저장",
            defaultextension=".pdf",
            filetypes=(("PDF 파일", "*.pdf"),),
            initialfile=f"{os.path.splitext(os.path.basename(self.input_file_path))[0]}_compressed.pdf"
        )

        if not output_path:
            return # 사용자가 저장을 취소한 경우

        quality = self.quality_var.get()
        
        # 압축 실행
        self.root.config(cursor="wait") # 마우스 커서를 대기 모양으로 변경
        self.root.update()

        success = compress_pdf(self.input_file_path, output_path, quality)
        
        self.root.config(cursor="") # 마우스 커서 원래대로
        
        if success:
            original_size = os.path.getsize(self.input_file_path) / (1024 * 1024) # MB
            compressed_size = os.path.getsize(output_path) / (1024 * 1024) # MB
            reduction = ((original_size - compressed_size) / original_size) * 100
            
            info_message = (
                f"압축이 완료되었습니다!\n\n"
                f"원본 크기: {original_size:.2f} MB\n"
                f"압축 후 크기: {compressed_size:.2f} MB\n"
                f"파일 크기 감소율: {reduction:.1f}%"
            )
            messagebox.showinfo("완료", info_message)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCompressorApp(root)
    root.mainloop()
