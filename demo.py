import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import smtplib
import email.utils
from email.mime.text import MIMEText

# 创建一个函数，用于更新Excel文件路径
def update_excel_path():
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if excel_file_path:
        # 在这里执行读取Excel文件的操作，你可以将下面的代码放入这个函数中
        # 读取Excel文件
        df = pd.read_excel(excel_file_path)
        # 获取学生姓名、科目、成绩和绩点列
        names = df['Unnamed: 2']
        courses = df['Unnamed: 3']
        grades = df['Unnamed: 5']
        gpa = df['Unnamed: 6']

         # 创建一个数组，用于存储每个学生的数据
        students_data = []

        current_student = None  # 用于跟踪当前学生的姓名

        for i in range(len(names)):
            student_name = names[i]
            course_name = courses[i]
            grade_value = grades[i]
            gpa_value = gpa[i]
            
            # 跳过第一行的列名
            if i == 0:
                continue
            
            # 如果遇到新的学生姓名，创建一个新学生的数据字典
            if student_name != current_student:
                if current_student is not None:
                    students_data.append(student_data)  # 将上一个学生的数据添加到数组
                current_student = student_name
                student_data = {'name': student_name, 'courses': [], 'grades': [], 'gpa': []}
            
            # 将课程和成绩添加到当前学生的数据中
            student_data['courses'].append(course_name)
            student_data['grades'].append(grade_value)
            student_data['gpa'].append(gpa_value)

        # 添加最后一个学生的数据
        if student_data:
            students_data.append(student_data)

        # 遍历学生数据数组并写入不同的txt文件
        for student_data in students_data:
            student_name = student_data['name']
            courses = student_data['courses']
            grades = student_data['grades']
            gpas = student_data['gpa']
            
            file_name = f'{student_name}.txt'
            
            with open(file_name, 'w', encoding='utf-8') as file:
                file.write(f'亲爱的{student_name}同学:\n')
                file.write('祝贺您顺利完成本学期的学习！教务处在此向您发送最新的成绩单。\n')
                for course, grade, gpa in zip(courses, grades, gpas):
                    file.write(f'{course}: 百分成绩：{grade} 绩点: {gpa}\n')
                file.write('希望您能够对自己的成绩感到满意，并继续保持努力和积极的学习态度。如果您在某些科目上没有达到预期的成绩，不要灰心，这也是学习过程中的一部分。我们鼓励您与您的任课教师或辅导员进行交流，他们将很乐意为您解答任何疑问并提供帮助。请记住，学习是一个持续不断的过程，我们相信您有能力克服困难并取得更大的进步。再次恭喜您，祝您学习进步、事业成功！教务处')
    
        # 更新路径显示
        path_label.config(text=f"Excel文件路径: {excel_file_path}")

def handle_input():
    user_input = entry.get()
    address = entry1.get()
    ans=user_input
    #print(f"你输入的内容是:"+ans)
    with open(ans+'.txt', 'r', encoding="utf-8") as file:
        data = file.read().rstrip()
    message = MIMEText(data)
    message['To'] = email.utils.formataddr(('接收者', address))
    message['From'] = email.utils.formataddr(('发送者', '2196055715@qq.com'))
    message['Subject'] = '我是邮件的标题'
    server = smtplib.SMTP_SSL('smtp.qq.com', 465)
    server.login('2196055715@qq.com', 'swjtyyurpobidjig')
    server.set_debuglevel(True)
    try:
        server.sendmail('2196055715@qq.com', [address], msg=message.as_string())
    finally:
        server.quit()


# 创建一个GUI窗口
root = tk.Tk()
root.title("选择Excel文件")
# 设置窗口大小
root.geometry("300x200")


# 选择文件按钮
select_button = tk.Button(root, text="选择Excel文件", command=update_excel_path)
select_button.pack(pady=10)

# 创建一个标签显示输入框的用途
lable2 = tk.Label(root, text="请输入姓名")
lable2.pack()

# 输入姓名
entry = tk.Entry(root)
entry.pack()

# 创建一个标签显示输入框的用途
lable3 = tk.Label(root, text="请输入邮箱")
lable3.pack()

#输入邮箱
entry1 = tk.Entry(root)
entry1.pack()

# 发送按钮
button = tk.Button(root, text="发送", command=handle_input)
button.pack()

# 创建一个标签显示Excel文件路径
path_label = tk.Label(root, text="Excel文件路径: 未选择")
path_label.pack()

# 启动GUI
root.mainloop()
