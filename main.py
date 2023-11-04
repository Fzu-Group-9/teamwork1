import os
import pandas as pd

# 获取当前脚本所在的文件夹
current_directory = os.path.dirname(__file__)

# 构建Excel文件的完整路径（假设Excel文件名为data.xlsx）
excel_file_path = os.path.join(current_directory, 'dan.xlsx')

# 读取Excel文件
df = pd.read_excel(excel_file_path)
#2,3,5
student_names = df['Unnamed: 2']
student_ke = df['Unnamed: 3']
student_chenji = df['Unnamed: 5']
selected_columns = df[['Unnamed: 2', 'Unnamed: 3', 'Unnamed: 5', 'Unnamed: 6']]

new_df = pd.DataFrame(selected_columns)

# 打印新的DataFrame
print(new_df)
#创建一个新字典
score_dict = {}
for index, row in df.iterrows():
    student_name = row['Unnamed: 2']
    subject = row['Unnamed: 3']
    score = row['Unnamed: 5']
    jidian = row['Unnamed: 6']

    if student_name in score_dict:
        # 如果姓名已存在于字典中，将成绩和科目添加到对应的列表中
        score_dict[student_name].append((subject, score, jidian))
    else:
        # 如果姓名不在字典中，创建一个新的列表并添加成绩和科目的元组
        score_dict[student_name] = [(subject, score,jidian )]

# 打印字典，其中键是学生姓名，值是成绩数组
print(score_dict)

