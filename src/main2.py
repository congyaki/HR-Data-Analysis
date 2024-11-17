import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from math import pi

# Thiết lập kết nối đến SQL Server

def connect_to_db():
    conn = pyodbc.connect(
        'DRIVER={SQL Server};'
        'SERVER=DESKTOP-AVPLSS5\SQLEXPRESS;'
        'DATABASE=HRManagement;'
        'UID=sa;'
        'PWD=123456@Aa;'
    )
    return conn

# Truy vấn dữ liệu và chuẩn bị dữ liệu
def get_employee_distribution_by_department(conn):
    query = """
    SELECT d.DepartmentName, COUNT(e.EmployeeId) AS EmployeeCount
    FROM Employees e
    JOIN Departments d ON e.DepartmentId = d.DepartmentId
    GROUP BY d.DepartmentName;
    """
    df = pd.read_sql(query, conn)
    return df

def get_employee_age_distribution(conn):
    query = """
    SELECT Age, COUNT(EmployeeId) AS EmployeeCount
    FROM Employees
    GROUP BY Age;
    """
    df = pd.read_sql(query, conn)
    return df

def get_performance_evaluations(conn, order_by='DESC', top_n=10):
    query = f"""
    SELECT 
        EmployeeId,
        FullName,
        AverageScore
    FROM EmployeePerformance
    ORDER BY AverageScore {order_by}
    OFFSET 0 ROWS FETCH NEXT {top_n} ROWS ONLY;
    """
    df = pd.read_sql(query, conn)
    return df

def get_training_programs(conn):
    query = """
    SELECT tp.ProgramName, COUNT(et.EmployeeId) AS ParticipationCount
    FROM TrainingPrograms tp
    JOIN EmployeeTraining et ON tp.ProgramId = et.ProgramId
    GROUP BY tp.ProgramName;
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
def plot_employee_distribution_by_department(df):
    plt.figure(figsize=(10, 6))
    sns.barplot(x='DepartmentName', y='EmployeeCount', data=df, palette='viridis')
    plt.title('Phân bố nhân sự theo phòng ban')
    plt.xlabel('Phòng ban')
    plt.ylabel('Số lượng nhân viên')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def plot_employee_age_distribution_bar(df):
    plt.figure(figsize=(10, 6))
    sns.barplot(x='Age', y='EmployeeCount', data=df, palette='coolwarm')
    plt.title('Phân bố độ tuổi nhân sự')
    plt.xlabel('Độ tuổi')
    plt.ylabel('Số lượng nhân viên')
    plt.show()

def plot_employee_seniority(df):
    plt.figure(figsize=(10, 6))
    sns.barplot(x='LevelName', y='EmployeeCount', data=df, palette='cubehelix')
    plt.title('Phân bố thâm niên nhân viên')
    plt.xlabel('Cấp độ Thâm niên')
    plt.ylabel('Số lượng nhân viên')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def plot_performance_bar(df):
    plt.figure(figsize=(10, 6))
    # Sắp xếp theo điểm đánh giá trung bình giảm dần
    df_sorted = df.sort_values('AverageScore', ascending=False)

    # Vẽ biểu đồ cột
    sns.barplot(x='FullName', y='AverageScore', data=df_sorted, palette='Blues_d')

    plt.title('Đánh giá hiệu suất nhân viên')
    plt.xlabel('Nhân viên')
    plt.ylabel('Điểm đánh giá trung bình')
    plt.xticks(rotation=45, ha='right')  # Xoay nhãn trục X cho dễ đọc
    plt.tight_layout()
    plt.show()




def plot_training_program_line(df):
    plt.figure(figsize=(10, 6))
    sns.lineplot(x='ProgramName', y='ParticipationCount', data=df, marker='o', sort=False, palette='husl')
    plt.title('Sự tham gia Chương trình Đào tạo')
    plt.xlabel('Chương trình Đào tạo')
    plt.ylabel('Số lượng Tham gia')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

# Cập nhật main
def main():
    conn = connect_to_db()
    print("Chọn biểu đồ để hiển thị:")
    print("1. Phân bố nhân sự theo phòng ban")
    print("2. Phân bố độ tuổi nhân sự (biểu đồ thanh ngang)")
    print("3. Đánh giá hiệu suất nhân viên (Bar Chart)")
    print("4. Sự tham gia chương trình đào tạo (Line Chart)")
    print("5. Thâm niên nhân viên")
    
    choice = input("Nhập lựa chọn (1-5): ")
    
    if choice == '1':
        df = get_employee_distribution_by_department(conn)
        plot_employee_distribution_by_department(df)
    elif choice == '2':
        df = get_employee_age_distribution(conn)
        plot_employee_age_distribution_bar(df)
    elif choice == '3':
        # Hỏi người dùng lựa chọn thứ tự sắp xếp và số lượng nhân viên
        order_choice = input("Chọn thứ tự sắp xếp (1: Tăng dần, 2: Giảm dần): ")
        order_by = 'ASC' if order_choice == '1' else 'DESC'
        top_n = int(input("Nhập số lượng nhân viên muốn hiển thị: "))

        df = get_performance_evaluations(conn, order_by, top_n)
        plot_performance_bar(df)
    elif choice == '4':
        df = get_training_programs(conn)
        plot_training_program_line(df)
    elif choice == '5':
        df = get_employee_seniority(conn)
        plot_employee_seniority(df)
    else:
        print("Lựa chọn không hợp lệ!")

    conn.close()

if __name__ == "__main__":
    main()
