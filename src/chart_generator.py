import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

# Cấu hình cho seaborn để làm đẹp biểu đồ
sns.set(style="whitegrid")

def plot_department_distribution(df):
    plt.figure(figsize=(10, 6))
    sns.barplot(x='EmployeeCount', y='DepartmentName', data=df, palette='Blues_d')
    plt.title('Phân bố nhân sự theo phòng ban', fontsize=16)
    plt.xlabel('Số lượng nhân viên')
    plt.ylabel('Phòng ban')
    plt.show()

def plot_age_distribution(df):
    plt.figure(figsize=(8, 6))
    sns.barplot(x='Age', y='Count', data=df, palette='Greens_d')
    plt.title('Phân bố độ tuổi nhân sự', fontsize=16)
    plt.xlabel('Độ tuổi')
    plt.ylabel('Số lượng nhân viên')
    plt.show()

def plot_performance_distribution(df):
    plt.figure(figsize=(10, 6))
    sns.barplot(x='EmployeeName', y='PerformanceScore', data=df, palette='Purples_d')
    plt.title('Đánh giá hiệu suất nhân viên', fontsize=16)
    plt.xlabel('Nhân viên')
    plt.ylabel('Điểm hiệu suất')
    plt.xticks(rotation=90)
    plt.show()
