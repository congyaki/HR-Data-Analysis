import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from tkinter import Tk, Label, Button, Entry, IntVar, StringVar, messagebox, ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# Kết nối cơ sở dữ liệu
def connect_to_db():
    conn = pyodbc.connect(
        'DRIVER={SQL Server};'
        'SERVER=DESKTOP-AVPLSS5\SQLEXPRESS;'
        'DATABASE=HRManagement;'
        'UID=sa;'
        'PWD=123456@Aa;'
    )
    return conn


# Các hàm truy vấn
def get_employee_distribution_by_department(conn, order_by='DESC', top_n=10):
    query = f"""
    SELECT TOP ({top_n}) d.DepartmentName, COUNT(e.EmployeeId) AS EmployeeCount
    FROM Employees e
    JOIN Departments d ON e.DepartmentId = d.DepartmentId
    GROUP BY d.DepartmentName
    ORDER BY EmployeeCount {order_by};
    """
    df = pd.read_sql(query, conn)
    return df


def get_employee_age_distribution(conn, age_from, age_to, order_by='DESC'):
    query = f"""
    SELECT Age, COUNT(EmployeeId) AS EmployeeCount
    FROM Employees
    WHERE Age BETWEEN {age_from} AND {age_to}
    GROUP BY Age
    ORDER BY EmployeeCount {order_by};
    """
    df = pd.read_sql(query, conn)
    print(df)

    return df



def get_performance_evaluations(conn, year=None, order_by='DESC', top_n=10):
    query = """
    SELECT 
        e.EmployeeId,
        e.FullName,
        AVG(pe.Score) AS AverageScore
    FROM 
        PerformanceEvaluations pe
    JOIN 
        Employees e ON pe.EmployeeId = e.EmployeeId
    """
    
    # Thêm điều kiện WHERE nếu người dùng chọn một năm cụ thể
    if year:
        query += f" WHERE YEAR(pe.EvaluationDate) = {year}"
    
    query += """
    GROUP BY 
        e.EmployeeId, e.FullName
    ORDER BY 
        AverageScore {order_by}
    OFFSET 0 ROWS FETCH NEXT {top_n} ROWS ONLY;
    """
    
    # Chèn các tham số vào truy vấn
    query = query.format(order_by=order_by, top_n=top_n)
    
    df = pd.read_sql(query, conn)
    return df





def get_training_programs(conn, top_n, order_by='DESC'):
    query = f"""
    SELECT tp.ProgramName, COUNT(et.EmployeeId) AS ParticipationCount
    FROM TrainingPrograms tp
    JOIN EmployeeTraining et ON tp.ProgramId = et.ProgramId
    GROUP BY tp.ProgramName
    ORDER BY ParticipationCount {order_by}
    OFFSET 0 ROWS FETCH NEXT {top_n} ROWS ONLY;
    """
    df = pd.read_sql(query, conn)
    return df


def get_seniority_level_descriptions(conn):
    query = """
    SELECT LevelName, Description
    FROM SeniorityLevels
    """
    df = pd.read_sql(query, conn)
    return df

def get_employee_seniority(conn):
    query = """
    SELECT s.LevelName, COUNT(es.EmployeeId) AS EmployeeCount
    FROM EmployeeSeniority es
    JOIN SeniorityLevels s ON es.SeniorityLevelId = s.SeniorityLevelId
    GROUP BY s.LevelName;
    """
    df = pd.read_sql(query, conn)
    return df


# Các hàm vẽ biểu đồ
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import MaxNLocator

def plot_employee_distribution_by_department(df):
    # Tạo kích thước biểu đồ
    plt.figure(figsize=(12, 8))  # Kích thước rộng hơn để chứa nhiều phòng ban

    # Cải thiện giao diện của seaborn
    sns.set(style="whitegrid")

    # Vẽ biểu đồ cột ngang
    ax = sns.barplot(x='EmployeeCount', y='DepartmentName', data=df, palette='viridis')

    # Thêm số liệu vào các cột
    for p in ax.patches:
        # Làm tròn số lượng nhân viên thành số nguyên
        ax.annotate(f'{int(p.get_width())}', 
                    (p.get_width(), p.get_y() + p.get_height() / 2.), 
                    ha='left', va='center', 
                    fontsize=12, color='white', 
                    xytext=(5, 0), textcoords='offset points')

    # Tiêu đề và nhãn trục
    plt.title('Phân bố nhân sự theo phòng ban', fontsize=16, weight='bold')
    plt.xlabel('Số lượng nhân viên', fontsize=14)
    plt.ylabel('Phòng ban', fontsize=14)

    # Đảm bảo trục x hiển thị số nguyên
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))

    # Cải thiện lưới và đường kẻ
    plt.grid(True, axis='x', linestyle='--', alpha=0.7)

    # Đảm bảo căn chỉnh phù hợp và không bị che khuất
    plt.tight_layout()  # Cải thiện việc căn chỉnh các phần tử của biểu đồ

    plt.show()




def plot_employee_age_distribution_bar(df, order_by='DESC'):
    # Sắp xếp dữ liệu theo EmployeeCount
    df = df.sort_values(by='EmployeeCount', ascending=(order_by == 'ASC'))
    
    # Lấy danh sách tuổi theo thứ tự xuất hiện trong DataFrame
    age_order = df['Age'].tolist()
    
    plt.figure(figsize=(12, 7))
    
    # Sử dụng màu sắc đẹp và dễ nhìn với bảng màu 'coolwarm'
    sns.barplot(x='Age', y='EmployeeCount', data=df, palette='coolwarm', order=age_order, width=0.8)

    # Thêm nhãn vào các cột để hiển thị số lượng nhân viên
    for p in plt.gca().patches:
        plt.gca().annotate(f'{int(p.get_height())}', 
                           (p.get_x() + p.get_width() / 2., p.get_height()), 
                           ha='center', va='center', 
                           fontsize=12, color='black', 
                           xytext=(0, 10), textcoords='offset points')
    
    # Tiêu đề và nhãn trục
    plt.title('Phân bố độ tuổi nhân sự', fontsize=18, weight='bold', color='#2E3A59')
    plt.xlabel('Độ tuổi', fontsize=14, color='#333333')
    plt.ylabel('Số lượng nhân viên', fontsize=14, color='#333333')

    # Cải thiện trục X để tránh nhãn bị chồng chéo
    plt.xticks(rotation=0, ha='right', fontsize=12, color='#555555')

    # Tùy chỉnh trục Y để bắt đầu từ 0 và hiển thị số nguyên
    plt.ylim(0, df['EmployeeCount'].max() + 10)
    plt.gca().yaxis.get_major_locator().set_params(integer=True)  # Đảm bảo trục Y là số nguyên

    # Thêm grid nhẹ nhàng cho biểu đồ dễ nhìn hơn
    plt.grid(True, axis='y', linestyle='--', alpha=0.7)

    # Tăng độ dày đường viền của cột và tạo hiệu ứng bóng
    for bar in plt.gca().patches:
        bar.set_edgecolor('black')
        bar.set_linewidth(1.5)
    
    # Hiển thị biểu đồ
    plt.tight_layout()
    plt.show()





def plot_performance_evaluations(df, order_by='DESC'):
    # Sắp xếp dữ liệu theo thứ tự do người dùng chọn
    df = df.sort_values(by='AverageScore', ascending=(order_by == 'ASC'))
    
    plt.figure(figsize=(12, 7))  # Kích thước lớn hơn cho biểu đồ đẹp hơn
    # Chuyển biểu đồ thành dạng ngang và dùng màu sắc đẹp với bảng màu
    sns.barplot(x='AverageScore', y='FullName', data=df, palette='Blues_d', edgecolor='black', linewidth=1.5)
    
    # Thêm nhãn cho điểm đánh giá ở cuối các cột bar
    for p in plt.gca().patches:
        plt.gca().annotate(f'{p.get_width():.2f}',  # Hiển thị số điểm với 2 chữ số thập phân
                           (p.get_width(), p.get_y() + p.get_height() / 2.), 
                           ha='left', va='center', 
                           fontsize=12, color='black', 
                           xytext=(5, 0), textcoords='offset points')

    # Tiêu đề và nhãn trục với màu sắc và font chữ đẹp hơn
    plt.title('Đánh giá hiệu suất nhân viên', fontsize=18, weight='bold', color='#2E3A59')
    plt.xlabel('Điểm đánh giá trung bình', fontsize=14, color='#333333')
    plt.ylabel('Nhân viên', fontsize=14, color='#333333')

    # Cải thiện trục X và Y để tránh nhãn bị chồng chéo và dễ nhìn hơn
    plt.xticks(rotation=0, fontsize=12, color='#555555')
    plt.yticks(fontsize=12, color='#555555')

    # Thêm grid nhẹ nhàng cho biểu đồ dễ nhìn hơn
    plt.grid(True, axis='x', linestyle='--', alpha=0.7)

    # Hiển thị biểu đồ
    plt.tight_layout()
    plt.show()



def plot_training_programs(df):
    plt.figure(figsize=(12, 8))  # Kích thước lớn hơn cho biểu đồ đẹp hơn
    # Sử dụng bảng màu 'coolwarm' và tăng độ dày đường viền của cột
    sns.barplot(x='ParticipationCount', y='ProgramName', data=df, palette='coolwarm', edgecolor='black', linewidth=1.5)
    
    # Thêm nhãn vào các cột để hiển thị số lượng tham gia
    for p in plt.gca().patches:
        plt.gca().annotate(f'{int(p.get_width())}',  # Hiển thị số lượng tham gia với số nguyên
                           (p.get_width(), p.get_y() + p.get_height() / 2.), 
                           ha='left', va='center', 
                           fontsize=12, color='black', 
                           xytext=(5, 0), textcoords='offset points')

    # Tiêu đề và nhãn trục với màu sắc và font chữ đẹp hơn
    plt.title('Sự tham gia chương trình đào tạo', fontsize=18, weight='bold', color='#2E3A59')
    plt.xlabel('Số lượng tham gia', fontsize=14, color='#333333')
    plt.ylabel('Chương trình đào tạo', fontsize=14, color='#333333')

    # Cải thiện trục X và Y để tránh nhãn bị chồng chéo và dễ nhìn hơn
    plt.xticks(rotation=0, fontsize=12, color='#555555')
    plt.yticks(fontsize=12, color='#555555')

    # Thêm grid nhẹ nhàng cho biểu đồ dễ nhìn hơn
    plt.grid(True, axis='x', linestyle='--', alpha=0.7)

    # Hiển thị biểu đồ
    plt.tight_layout()
    plt.show()



def plot_employee_seniority(df):
    plt.figure(figsize=(12, 7))  # Tăng kích thước của biểu đồ
    sns.set(style="whitegrid")  # Thiết lập kiểu nền sáng và lưới trắng
    
    # Vẽ biểu đồ với gradient màu sắc
    ax = sns.barplot(x='LevelName', y='EmployeeCount', data=df, palette='viridis')  
    
    # Thêm số liệu vào các cột của biểu đồ
    for p in ax.patches:
        ax.annotate(f'{p.get_height()}', 
                    (p.get_x() + p.get_width() / 2., p.get_height()), 
                    ha='center', va='center', 
                    fontsize=12, color='white', 
                    xytext=(0, 9), textcoords='offset points')

    # Cải thiện tiêu đề và nhãn trục
    plt.title('Phân bố thâm niên nhân viên', fontsize=16, weight='bold')
    plt.xlabel('Cấp độ thâm niên', fontsize=14)
    plt.ylabel('Số lượng nhân viên', fontsize=14)
    
    # Đặt các nhãn trục X theo chiều ngang
    plt.xticks(rotation=0, fontsize=12)  
    
    # Điều chỉnh lưới
    plt.grid(True, axis='y', linestyle='--', alpha=0.7)

    plt.tight_layout()  # Cải thiện việc căn chỉnh các phần tử của biểu đồ
    plt.show()


# Hàm xử lý tab
def show_employee_distribution_by_department():
    try:
        top_n = int(department_top_n_var.get())
        order_by = department_order_var.get()
        conn = connect_to_db()
        df = get_employee_distribution_by_department(conn, order_by, top_n)
        conn.close()
        plot_employee_distribution_by_department(df)
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))



def show_employee_age_distribution():
    try:
        # Lấy giá trị từ giao diện
        age_from = int(age_from_var.get())
        age_to = int(age_to_var.get())
        order_by = age_order_var.get()

        # Kiểm tra điều kiện hợp lệ
        if age_from > age_to:
            messagebox.showerror("Lỗi", "Độ tuổi bắt đầu phải nhỏ hơn hoặc bằng độ tuổi kết thúc.")
            return

        # Kết nối và thực thi truy vấn
        conn = connect_to_db()
        df = get_employee_age_distribution(conn, age_from, age_to, order_by)
        conn.close()

        # Kiểm tra nếu không có dữ liệu trả về
        if df.empty:
            messagebox.showinfo("Thông báo", "Không có dữ liệu phù hợp với điều kiện.")
            return

        # Hiển thị biểu đồ, truyền order_by để vẽ đúng thứ tự
        plot_employee_age_distribution_bar(df, order_by)

    except ValueError:
        messagebox.showerror("Lỗi", "Vui lòng nhập số nguyên hợp lệ cho độ tuổi.")
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


def show_performance_evaluations_with_order():
    try:
        year = performance_year_var.get()
        year = int(year) if year.isdigit() else None  # Nếu không nhập số, coi như tất cả các năm
        
        order_by = performance_order_var.get()
        top_n = int(performance_top_n_var.get())
        
        # Kiểm tra số lượng nhân viên nhập vào
        if top_n > 30:
            # Hiển thị thông báo lỗi và yêu cầu nhập lại
            messagebox.showwarning("Lỗi", "Số lượng nhân viên phải nhỏ hơn hoặc bằng 30. Vui lòng nhập lại.")
            performance_top_n_var.set(10)  # Đặt lại giá trị mặc định cho Entry (10 nhân viên)
            return  # Dừng hàm nếu số lượng vượt quá 30
        
        # Kết nối tới cơ sở dữ liệu và lấy dữ liệu
        conn = connect_to_db()
        df = get_performance_evaluations(conn, year, order_by, top_n)
        conn.close()
        
        if df.empty:
            messagebox.showinfo("Thông báo", "Không có dữ liệu phù hợp.")
        else:
            plot_performance_evaluations(df, order_by)
    except ValueError:
        messagebox.showerror("Lỗi", "Vui lòng nhập năm hoặc số lượng nhân viên hợp lệ.")
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))




def show_training_programs():
    try:
        top_n = int(training_top_n_var.get())
        order_by = training_order_var.get()  # Lấy giá trị sắp xếp
        conn = connect_to_db()
        df = get_training_programs(conn, top_n, order_by)
        conn.close()
        plot_training_programs(df)
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))



def show_employee_seniority():
    try:
        conn = connect_to_db()
        df = get_employee_seniority(conn)
        conn.close()
        plot_employee_seniority(df)
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


# Tạo giao diện Tkinter
root = Tk()
root.title("Phân tích dữ liệu nhân sự")
root.geometry("400x400")

notebook = ttk.Notebook(root)

# Tab 1: Phân bố theo phòng ban
tab_department = ttk.Frame(notebook)
notebook.add(tab_department, text="Phân bố nhân sự theo phòng ban")

Label(tab_department, text="Nhập số lượng phòng ban muốn hiển thị:").pack(pady=5)
department_top_n_var = IntVar()
Entry(tab_department, textvariable=department_top_n_var).pack(pady=5)

# Thứ tự sắp xếp
Label(tab_department, text="Chọn thứ tự sắp xếp:").pack(pady=5)
department_order_var = StringVar(value='DESC')
ttk.Radiobutton(tab_department, text="Tăng dần", variable=department_order_var, value='ASC').pack()
ttk.Radiobutton(tab_department, text="Giảm dần", variable=department_order_var, value='DESC').pack()

Button(tab_department, text="Hiển thị biểu đồ", command=show_employee_distribution_by_department).pack(pady=20)


# Tab 2: Phân bố theo độ tuổi
tab_age = ttk.Frame(notebook)
notebook.add(tab_age, text="Phân bố nhân sự theo độ tuổi")

Label(tab_age, text="Nhập độ tuổi bắt đầu:").pack(pady=5)
age_from_var = IntVar(value=20)  # Giá trị mặc định
Entry(tab_age, textvariable=age_from_var).pack(pady=5)

Label(tab_age, text="Nhập độ tuổi kết thúc:").pack(pady=5)
age_to_var = IntVar(value=50)  # Giá trị mặc định
Entry(tab_age, textvariable=age_to_var).pack(pady=5)

Label(tab_age, text="Chọn thứ tự sắp xếp:").pack(pady=5)
age_order_var = StringVar(value='DESC')
ttk.Radiobutton(tab_age, text="Tăng dần", variable=age_order_var, value='ASC').pack()
ttk.Radiobutton(tab_age, text="Giảm dần", variable=age_order_var, value='DESC').pack()

Button(tab_age, text="Hiển thị biểu đồ", command=show_employee_age_distribution).pack(pady=20)


# Tab 3: Đánh giá hiệu suất nhân viên
tab_performance = ttk.Frame(notebook)
notebook.add(tab_performance, text="Đánh giá hiệu suất của nhân viên")

# Thứ tự sắp xếp
Label(tab_performance, text="Chọn thứ tự sắp xếp:").pack(pady=5)
performance_order_var = StringVar(value='DESC')
ttk.Radiobutton(tab_performance, text="Tăng dần", variable=performance_order_var, value='ASC').pack()
ttk.Radiobutton(tab_performance, text="Giảm dần", variable=performance_order_var, value='DESC').pack()

# Số lượng nhân viên
Label(tab_performance, text="Nhập số lượng nhân viên muốn hiển thị:").pack(pady=5)
performance_top_n_var = IntVar(value=10)
Entry(tab_performance, textvariable=performance_top_n_var).pack(pady=5)

# Chọn năm hoặc tất cả các năm
Label(tab_performance, text="Nhập năm (để trống để xem tất cả các năm):").pack(pady=5)
performance_year_var = StringVar()  # Sử dụng StringVar để cho phép nhập trống
Entry(tab_performance, textvariable=performance_year_var).pack(pady=5)

Button(tab_performance, text="Hiển thị biểu đồ", command=show_performance_evaluations_with_order).pack(pady=20)

# Tab 4: Tham gia chương trình đào tạo
tab_training = ttk.Frame(notebook)
notebook.add(tab_training, text="Đào tạo và phát triển")

# Nhập số lượng chương trình muốn hiển thị
Label(tab_training, text="Nhập số lượng CTĐT muốn hiển thị:").pack(pady=5)
training_top_n_var = IntVar()
Entry(tab_training, textvariable=training_top_n_var).pack(pady=5)

# Thứ tự sắp xếp
Label(tab_training, text="Chọn thứ tự sắp xếp theo số lượng tham gia:").pack(pady=5)
training_order_var = StringVar(value='DESC')  # Mặc định là giảm dần
ttk.Radiobutton(tab_training, text="Tăng dần", variable=training_order_var, value='ASC').pack()
ttk.Radiobutton(tab_training, text="Giảm dần", variable=training_order_var, value='DESC').pack()

# Hiển thị biểu đồ
Button(tab_training, text="Hiển thị biểu đồ", command=show_training_programs).pack(pady=20)

# Tab 5: Thâm niên nhân viên
tab_seniority = ttk.Frame(notebook)
notebook.add(tab_seniority, text="Số lượng nhân viên theo cấp độ thâm niên")

Button(tab_seniority, text="Hiển thị biểu đồ", command=show_employee_seniority).pack(pady=20)

# Hiển thị các tab
notebook.pack(expand=True, fill='both')

root.mainloop()
