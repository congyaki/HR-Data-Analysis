from tkinter import *
from database import query_department_distribution, query_age_distribution
from chart_generator import plot_department_distribution, plot_age_distribution

def show_department_chart():
    df = query_department_distribution()
    if df is not None:
        plot_department_distribution(df)

def show_age_chart():
    df = query_age_distribution()
    if df is not None:
        plot_age_distribution(df)

def create_gui():
    window = Tk()
    window.title("Quản lý nhân sự")
    window.geometry("400x300")

    label = Label(window, text="Chọn biểu đồ cần hiển thị:", font=("Arial", 14))
    label.pack(pady=20)

    button1 = Button(window, text="Biểu đồ phân bố nhân sự theo phòng ban", width=30, command=show_department_chart)
    button1.pack(pady=10)

    button2 = Button(window, text="Biểu đồ phân bố độ tuổi nhân sự", width=30, command=show_age_chart)
    button2.pack(pady=10)

    window.mainloop()

if __name__ == "__main__":
    create_gui()
