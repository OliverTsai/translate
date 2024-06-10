import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from googletrans import Translator

# 初始化翻譯器
translator = Translator()

def browse_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        # 去除檔案名稱中的空格
        new_file_path = file_path.replace(' ', '_')
        
        if new_file_path != file_path:
            # 重新命名檔案
            try:
                os.rename(file_path, new_file_path)
                file_path = new_file_path
                file_label.config(text=new_file_path)
                messagebox.showinfo("檔案已選擇", f"檔案已重新命名: {new_file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"重新命名檔案失敗: {e}")
        else:
            file_label.config(text=new_file_path)
            messagebox.showinfo("檔案已選擇", f"檔案路徑: {new_file_path}")

def translate_cell(cell_value):
    if pd.notnull(cell_value):
        try:
            translated = translator.translate(str(cell_value), dest='zh-cn')
            print("翻譯結果：" + translated)
            return translated.text
        except Exception as e:
            print(f"翻譯失敗: {e}")
            print("文件內容："+cell_value)
            return cell_value
    return cell_value

def output():
    global file_path
    if not file_path:
        messagebox.showwarning("警告", "請先選擇檔案")
        return

    number = number_entry.get()
    name = name_entry.get()
    
    if not number.isdigit():
        messagebox.showwarning("警告", "請輸入有效的數字")
        return

    if not name:
        messagebox.showwarning("警告", "請輸入有效的檔案名稱")
        return

    number = int(number)
    
    # 讀取 Excel 文件
    df = pd.read_excel(file_path)
    
    if number < 1 or number > len(df.columns):
        messagebox.showwarning("警告", "請輸入有效的欄位數字")
        return
    
    # 翻譯指定欄位
    col = df.columns[number - 1]
    list_date = []
    count = 0
    for raw in df[col]:
        list_date.append(translate_cell(raw))
        count = count+1
        if count>10:
            break
    df[col] = list_date
    
    # 將翻譯結果保存到另一個 Excel 文件中
    output_file_path = f'{name}.xlsx'
    df.to_excel(output_file_path, index=False)
    
    messagebox.showinfo("輸出結果", f"翻譯完成並保存到 {output_file_path}")

# 初始化 Tkinter 應用程式
app = tk.Tk()
app.title("Tkinter 翻譯軟件")

# 設定視窗大小和位置
window_width = 500
window_height = 300

screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()

center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)

app.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

# 檔案選擇部分
browse_button = tk.Button(app, text="讀取檔案路徑", command=browse_file)
browse_button.pack(pady=10)

file_label = tk.Label(app, text="", wraplength=400)
file_label.pack(pady=10)

# 輸入數字部分
tk.Label(app, text="輸入要翻譯的欄位數字:").pack(pady=5)
number_entry = tk.Entry(app)
number_entry.pack(pady=5)

# 輸入名字部分
tk.Label(app, text="輸入檔案名稱:").pack(pady=5)
name_entry = tk.Entry(app)
name_entry.pack(pady=5)

# 輸出按鈕
output_button = tk.Button(app, text="輸出結果", command=output)
output_button.pack(pady=20)

file_path = ""
app.mainloop()