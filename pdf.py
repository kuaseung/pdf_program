import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pikepdf
import os
import subprocess
import shutil

# ------------------ 유틸 함수 ------------------
def format_size(bytes_size):
    for unit in ["B", "KB", "MB", "GB"]:
        if bytes_size < 1024:
            return f"{bytes_size:.2f} {unit}"
        bytes_size /= 1024
    return f"{bytes_size:.2f} TB"

def find_ghostscript():
    for name in ["gs", "gswin64c", "gswin32c"]:
        path = shutil.which(name)
        if path:
            return path
    return None

def ask_ghostscript_path():
    path = simpledialog.askstring("Ghostscript 경로", 
        "Ghostscript 실행 파일 경로를 입력해주세요 (예: C:\\Program Files\\gs\\gs10.01.3\\bin\\gswin64c.exe):")
    if path and os.path.exists(path):
        return path
    messagebox.showerror("에러", "유효하지 않은 경로입니다.")
    return None

# ------------------ PDF 압축 함수 ------------------
def compress_pdf_pikepdf(input_file, quality):
    try:
        pdf = pikepdf.open(input_file)
        for page in pdf.pages:
            for key, value in list(page.images.items()):
                try:
                    img = pikepdf.PdfImage(value)
                    page.images[key] = img.as_jpeg(quality=quality)
                except Exception:
                    continue
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_compressed.pdf"
        pdf.save(output_file)
        pdf.close()
        return output_file
    except Exception as e:
        print(f"pikepdf 압축 실패: {input_file}, {e}")
        return None

def compress_pdf_gs(input_file, gs_path, quality="ebook"):
    base, ext = os.path.splitext(input_file)
    output_file = f"{base}_compressed.pdf"
    cmd = [
        gs_path,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{quality}",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-sOutputFile={output_file}",
        input_file
    ]
    try:
        subprocess.run(cmd, check=True)
        return output_file
    except subprocess.CalledProcessError as e:
        print(f"Ghostscript 압축 실패: {input_file}, {e}")
        return None

# ------------------ GUI 이벤트 ------------------
def select_files():
    files = filedialog.askopenfilenames(title="압축할 PDF 파일 선택", filetypes=[("PDF Files", "*.pdf")])
    if files:
        listbox_files.delete(0, tk.END)
        for f in files:
            listbox_files.insert(tk.END, f)

def compress_pdfs():
    files = listbox_files.get(0, tk.END)
    if not files:
        messagebox.showerror("에러", "PDF 파일을 선택해주세요.")
        return

    method = compress_method.get()
    quality = slider_quality.get()
    result_text.delete("1.0", tk.END)

    success, fail = 0, 0

    # Ghostscript 경로
    gs_path = None
    if method == "gs":
        gs_path = find_ghostscript()
        if not gs_path:
            gs_path = ask_ghostscript_path()
        if not gs_path:
            messagebox.showerror("에러", "Ghostscript 경로를 찾을 수 없습니다.")
            return

    for input_file in files:
        original_size = os.path.getsize(input_file)

        if method == "pikepdf":
            output_file = compress_pdf_pikepdf(input_file, quality)
        else:
            output_file = compress_pdf_gs(input_file, gs_path, "ebook")

        if output_file and os.path.exists(output_file):
            compressed_size = os.path.getsize(output_file)
            compression_ratio = (original_size - compressed_size) / original_size * 100  # 압축률

            result_text.insert(tk.END,
                f"파일: {os.path.basename(input_file)}\n"
                f" - 원본 크기: {format_size(original_size)}\n"
                f" - 압축 크기: {format_size(compressed_size)}\n"
                f" - 압축률: {compression_ratio:.1f}%\n\n"
            )
            success += 1
        else:
            result_text.insert(tk.END, f"❌ 실패: {os.path.basename(input_file)}\n\n")
            fail += 1

    messagebox.showinfo("완료", f"압축 완료!\n성공: {success}개\n실패: {fail}개")

# ------------------ GUI ------------------
root = tk.Tk()
root.title("PDF 일괄 압축기")
root.minsize(650, 550)

# 파일 선택
frame = tk.Frame(root)
frame.pack(pady=10, fill="x")
btn_browse = tk.Button(frame, text="여러 개 파일 선택", command=select_files)
btn_browse.pack(side="left", padx=5)

# 파일 리스트
listbox_files = tk.Listbox(root, width=90, height=8, selectmode=tk.MULTIPLE)
listbox_files.pack(pady=10, padx=10)

# 압축 방식 선택
frame_method = tk.LabelFrame(root, text="압축 방식 선택", padx=10, pady=10)
frame_method.pack(pady=10, fill="x", padx=10)
compress_method = tk.StringVar(value="pikepdf")
tk.Radiobutton(frame_method, text="빠른 압축 (pikepdf)", variable=compress_method, value="pikepdf").pack(anchor="w")
tk.Radiobutton(frame_method, text="고압축 (Ghostscript)", variable=compress_method, value="gs").pack(anchor="w")

# 이미지 품질 슬라이더 (pikepdf 전용)
frame_slider = tk.LabelFrame(root, text="이미지 품질 (%) (pikepdf 전용)", padx=10, pady=10)
frame_slider.pack(pady=10, fill="x", padx=10)
slider_quality = tk.Scale(frame_slider, from_=10, to=100, orient="horizontal")
slider_quality.set(70)
slider_quality.pack(fill="x")

# 실행 버튼
btn_compress = tk.Button(root, text="📂 선택한 PDF 모두 압축하기", command=compress_pdfs, width=35, bg="lightblue")
btn_compress.pack(pady=10)

# 결과 출력창
frame_result = tk.LabelFrame(root, text="압축 결과", padx=10, pady=10)
frame_result.pack(pady=10, fill="both", expand=True, padx=10)
result_text = tk.Text(frame_result, wrap="word", height=12)
result_text.pack(fill="both", expand=True)

root.mainloop()
