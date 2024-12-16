import tkinter as tk
from tkinter import messagebox, ttk
import csv
import pandas as pd
from datetime import datetime

file_csv = "nhanvien.csv"

root = tk.Tk()
root.title("Quản lý nhân viên")
root.geometry("600x500")

def du_lieu():
    data = {
        "Mã": ma.get(),
        "Tên": ten.get(),
        "Ngày sinh": ngaysinh.get(),
        "Giới tính": gender_var.get(),
        "Số CMND": cmnd.get(),
        "Nơi cấp": noicap.get(),
        "Đơn vị": donvi.get(),
        "Chức danh": chucdanh.get(),
        "Ngày cấp": ngaycap.get()
    }

    if not all(data.values()):
        messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ thông tin.")
        return

    try:
        with open(file_csv, mode="a", newline="", encoding="utf-8") as file:
            write = csv.DictWriter(file, fieldnames=data.keys())
            if file.tell() == 0:
                write.writeheader()
            write.writerow(data)
        messagebox.showinfo("Thành công", "Dữ liệu đã được lưu.")
        o_nhap()
    except Exception as e:
        messagebox.showerror("Lỗi", f"Lỗi khi lưu dữ liệu: {e}")

def o_nhap():
    ma.delete(0, tk.END)
    ten.delete(0, tk.END)
    ngaysinh.delete(0, tk.END)
    cmnd.delete(0, tk.END)
    noicap.delete(0, tk.END)
    donvi.delete(0, tk.END)
    chucdanh.delete(0, tk.END)
    ngaycap.delete(0, tk.END)
    gender_var.set("")

def sinh_nhat():
    try:
        today = datetime.now().strftime("%d/%m/%Y")[0:5]
        birthday = []
        df = pd.read_csv(file_csv)
        for index, row in df.iterrows():
            if str(row["Ngày sinh"])[0:5] == today:
                birthday.append(row["Tên"])

        if birthday:
            messagebox.showinfo("Sinh nhật hôm nay",  "\n".join(birthday))
        else:
            messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")

    except Exception as e:
        messagebox.showerror("Lỗi", f"Lỗi khi kiểm tra sinh nhật: {e}")

def ghi_excel():
    try:
        df = pd.read_csv(file_csv)
        df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y", errors="coerce")
        df = df.sort_values(by="Ngày sinh", ascending=True)

        excel_file = "danhsach_nhanvien.xlsx"
        df.to_excel(excel_file, index=False)
        messagebox.showinfo("Thành công", f"Dữ liệu đã được xuất ra file {excel_file}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Lỗi khi xuất file Excel: {e}")

gender_var = tk.StringVar()

tk.Label(root, text="Mã nhân viên").grid(row=0, column=0, padx=10, pady=5)
ma = tk.Entry(root)
ma.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Tên").grid(row=1, column=0, padx=10, pady=5)
ten = tk.Entry(root)
ten.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Ngày sinh (DD/MM/YYYY)").grid(row=2, column=0, padx=10, pady=5)
ngaysinh = tk.Entry(root)
ngaysinh.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Giới tính").grid(row=3, column=0, padx=10, pady=5)
tk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam").grid(row=3, column=1, sticky="w")
tk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ").grid(row=3, column=1, sticky="e")

tk.Label(root, text="Số CMND").grid(row=4, column=0, padx=10, pady=5)
cmnd = tk.Entry(root)
cmnd.grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Nơi cấp").grid(row=5, column=0, padx=10, pady=5)
noicap = tk.Entry(root)
noicap.grid(row=5, column=1, padx=10, pady=5)

tk.Label(root, text="Đơn vị").grid(row=6, column=0, padx=10, pady=5)
donvi = tk.Entry(root)
donvi.grid(row=6, column=1, padx=10, pady=5)

tk.Label(root, text="Chức danh").grid(row=7, column=0, padx=10, pady=5)
chucdanh = tk.Entry(root)
chucdanh.grid(row=7, column=1, padx=10, pady=5)

tk.Label(root, text="Ngày cấp (DD/MM/YYYY)").grid(row=8, column=0, padx=10, pady=5)
ngaycap = tk.Entry(root)
ngaycap.grid(row=8, column=1, padx=10, pady=5)

nut_luu = tk.Button(root, text="Lưu dữ liệu", command=du_lieu, bg="lightgreen")
nut_luu.grid(row=9, column=0, pady=10)

nut_sinh_nhat = tk.Button(root, text="Sinh nhật hôm nay", command=sinh_nhat, bg="lightblue")
nut_sinh_nhat.grid(row=9, column=1, pady=10)

nut_xuat_file = tk.Button(root, text="Xuất toàn bộ danh sách", command=ghi_excel, bg="lightyellow")
nut_xuat_file.grid(row=10, column=0, columnspan=2, pady=10)

root.mainloop()