import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
from tkinter import simpledialog

class ExcelProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("处理Excel表格")

        # 创建选项卡
        self.tabControl = ttk.Notebook(self.root)

        # 创建数据处理标签页
        self.tab_process = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab_process, text="数据处理")
        self.tabControl.pack(expand=1, fill="both")

        # 创建后台管理标签页
        self.tab_manage = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab_manage, text="后台管理")
        self.tabControl.pack(expand=1, fill="both")

        # 创建表格A文件路径输入框
        self.label_path_A = tk.Label(self.tab_process, text="选择表格A：")
        self.label_path_A.grid(row=0, column=0)
        self.entry_path_A = tk.Entry(self.tab_process, width=50)
        self.entry_path_A.grid(row=0, column=1)
        self.button_browse_A = tk.Button(self.tab_process, text="浏览", command=self.browse_button_A)
        self.button_browse_A.grid(row=0, column=2)

        # 创建表格B文件路径输入框
        self.label_path_B = tk.Label(self.tab_process, text="选择表格B：")
        self.label_path_B.grid(row=1, column=0)
        self.entry_path_B = tk.Entry(self.tab_process, width=50)
        self.entry_path_B.grid(row=1, column=1)
        self.button_browse_B = tk.Button(self.tab_process, text="浏览", command=self.browse_button_B)
        self.button_browse_B.grid(row=1, column=2)

        # 创建项目编号选择框
        self.label_project_number = tk.Label(self.tab_process, text="项目编号：")
        self.label_project_number.grid(row=2, column=0)
        self.project_number_var = tk.StringVar(value="")
        self.entry_project_number = tk.Entry(self.tab_process, textvariable=self.project_number_var, state='readonly')
        self.entry_project_number.grid(row=2, column=1)
        self.button_select_project = tk.Button(self.tab_process, text="选择项目编号", command=self.select_project)
        self.button_select_project.grid(row=2, column=2)

        # 创建相关人选择框
        self.label_related_person = tk.Label(self.tab_process, text="相关人：")
        self.label_related_person.grid(row=3, column=0)
        self.entry_related_person = tk.Entry(self.tab_process, state='readonly', width=50)
        self.entry_related_person.grid(row=3, column=1)
        self.button_select_users = tk.Button(self.tab_process, text="选择相关人", command=self.select_users)
        self.button_select_users.grid(row=3, column=2)

        # 创建承接部门选择框
        self.label_department_name = tk.Label(self.tab_process, text="承接部门名称：")
        self.label_department_name.grid(row=4, column=0)
        self.department_name_var = tk.StringVar(value="")
        self.entry_department_name = tk.Entry(self.tab_process, textvariable=self.department_name_var, state='readonly')
        self.entry_department_name.grid(row=4, column=1)
        self.button_select_department = tk.Button(self.tab_process, text="选择承接部门", command=self.select_department)
        self.button_select_department.grid(row=4, column=2)

        # 创建处理按钮
        self.button_process = tk.Button(self.tab_process, text="处理", command=self.process_button)
        self.button_process.grid(row=5, column=1)

        # 创建用户列表
        self.user_list = []

        # 创建用户列表框
        self.user_listbox = tk.Listbox(self.tab_manage)
        self.user_listbox.pack(side="left")

        # 创建添加用户按钮
        self.button_add_user = tk.Button(self.tab_manage, text="添加用户", command=self.add_user)
        self.button_add_user.pack(side="left")

        # 创建承接部门列表
        self.department_list = []

        # 创建承接部门列表框
        self.department_listbox = tk.Listbox(self.tab_manage)
        self.department_listbox.pack(side="left")

        # 创建添加承接部门按钮
        self.button_add_department = tk.Button(self.tab_manage, text="添加承接部门", command=self.add_department)
        self.button_add_department.pack(side="left")

        # 创建项目编号列表
        self.project_list = []

        # 创建项目编号列表框
        self.project_listbox = tk.Listbox(self.tab_manage)
        self.project_listbox.pack(side="left")

        # 创建添加项目编号按钮
        self.button_add_project = tk.Button(self.tab_manage, text="添加项目编号", command=self.add_project)
        self.button_add_project.pack(side="left")

    def browse_button_A(self):
        filename = filedialog.askopenfilename()
        self.entry_path_A.delete(0, tk.END)
        self.entry_path_A.insert(0, filename)

    def browse_button_B(self):
        filename = filedialog.askopenfilename()
        self.entry_path_B.delete(0, tk.END)
        self.entry_path_B.insert(0, filename)

    def process_button(self):
        input_file_A = self.entry_path_A.get()
        input_file_B = self.entry_path_B.get()
        project_number = self.project_number_var.get()
        department_name = self.department_name_var.get()

        # 获取用户选择的相关人
        related_person = self.entry_related_person.get()

        # 获取表格A所在的文件夹路径
        input_folder_A = os.path.dirname(input_file_A)

        self.process_excel(input_file_A, input_file_B, input_folder_A, project_number, related_person, department_name)

    def process_excel(self, input_file_A, input_file_B, output_folder, project_number, related_person, department_name):
        # 读取表格A和表格B的第一个sheet页
        df_A = pd.read_excel(input_file_A)
        df_B = pd.read_excel(input_file_B, sheet_name=0)  # 读取第一个sheet页

        # 根据承接部门名称和状态筛选表格A的数据
        filtered_df_A = df_A[(df_A['承接部门'] == department_name) & (df_A['状态'].isin(['新提交', '采购中']))]

        # 如果筛选后的数据为空，提示用户并退出程序
        if filtered_df_A.empty:
            print("没有符合条件的数据，请检查承接部门名称和状态列的内容。")
            return

        # 根据筛选后的数据生成任务名称列
        filtered_df_A['任务名称'] = filtered_df_A['售前编号'] + '-' + filtered_df_A['售前名称']

        # 将指派人列重命名为负责人
        filtered_df_A.rename(columns={'指派人': '负责人'}, inplace=True)

        # 只保留任务名称、负责人和状态列
        df_new = filtered_df_A[['任务名称', '负责人']]

        # 添加空的任务描述列
        df_new['任务描述'] = ''

        # 添加用户提供的项目编号和相关人信息
        df_new['项目编号'] = project_number
        df_new['相关人'] = related_person

        # 调整列顺序
        df_new = df_new[['项目编号', '任务名称', '负责人', '相关人', '任务描述']]

        # 生成新表格文件名，包含项目编号和时间戳
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        output_file = f"OA导入_{project_number}_{timestamp}.xlsx"
        output_path = os.path.join(output_folder, output_file)

        # 将处理后的数据写入新的Excel表格
        with pd.ExcelWriter(output_path) as writer:
            df_new.to_excel(writer, index=False, sheet_name='Sheet1')  # 写入第一个sheet页

            # 复制表格B的第二个sheet页到表格C的第二个sheet页中
            wb_B = pd.ExcelFile(input_file_B)
            sheets_B = wb_B.sheet_names
            for sheet_name in sheets_B:
                if sheet_name != 'Sheet1':
                    df_B_sheet2 = pd.read_excel(input_file_B, sheet_name=sheet_name)
                    df_B_sheet2.to_excel(writer, index=False, sheet_name=sheet_name)

        print("处理完成，已生成新的Excel表格:", output_path)

    def add_user(self):
        user = simpledialog.askstring("添加用户", "请输入用户名：")
        if user:
            self.user_list.append(user)
            self.update_user_listbox()

    def add_department(self):
        department = simpledialog.askstring("添加承接部门", "请输入承接部门名称：")
        if department:
            self.department_list.append(department)
            self.update_department_listbox()

    def add_project(self):
        project = simpledialog.askstring("添加项目编号", "请输入项目编号：")
        if project:
            self.project_list.append(project)
            self.update_project_listbox()

    def update_user_listbox(self):
        self.user_listbox.delete(0, tk.END)
        for user in self.user_list:
            self.user_listbox.insert(tk.END, user)

    def update_department_listbox(self):
        self.department_listbox.delete(0, tk.END)
        for department in self.department_list:
            self.department_listbox.insert(tk.END, department)

    def update_project_listbox(self):
        self.project_listbox.delete(0, tk.END)
        for project in self.project_list:
            self.project_listbox.insert(tk.END, project)

    def select_users(self):
        # 创建用户选择框
        self.user_frame = ttk.LabelFrame(self.tab_process, text="选择相关人")
        self.user_frame.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        # 创建用户列表框
        self.user_listbox_selection = tk.Listbox(self.user_frame, selectmode=tk.MULTIPLE)
        self.user_listbox_selection.pack(side="left")

        # 将已添加的用户显示在列表框中
        for user in self.user_list:
            self.user_listbox_selection.insert(tk.END, user)

        # 创建添加用户的文本框和按钮
        self.entry_add_user = tk.Entry(self.user_frame, width=30)
        self.entry_add_user.pack(side="left")
        self.entry_add_user.focus_set()
        self.entry_add_user.bind("<Return>", lambda event: self.add_user_popup())
        self.button_add_user_popup = tk.Button(self.user_frame, text="添加用户", command=self.add_user_popup)
        self.button_add_user_popup.pack(side="left")

        # 创建确认按钮
        button_confirm = tk.Button(self.user_frame, text="确认", command=self.confirm_selection)
        button_confirm.pack(side="left")

    def add_user_popup(self):
        user = self.entry_add_user.get()
        if user:
            self.user_list.append(user)
            self.user_listbox_selection.insert(tk.END, user)
            self.update_user_listbox()
            self.entry_add_user.delete(0, tk.END)

    def select_department(self):
        # 创建承接部门选择框
        self.department_frame = ttk.LabelFrame(self.tab_process, text="选择承接部门")
        self.department_frame.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        # 创建承接部门列表框
        self.department_listbox_selection = tk.Listbox(self.department_frame)
        self.department_listbox_selection.pack(side="left")

        # 将已添加的承接部门显示在列表框中
        for department in self.department_list:
            self.department_listbox_selection.insert(tk.END, department)

        # 创建添加承接部门的文本框和按钮
        self.entry_add_department = tk.Entry(self.department_frame, width=30)
        self.entry_add_department.pack(side="left")
        self.entry_add_department.focus_set()
        self.entry_add_department.bind("<Return>", lambda event: self.add_department_popup())
        self.button_add_department_popup = tk.Button(self.department_frame, text="添加承接部门", command=self.add_department_popup)
        self.button_add_department_popup.pack(side="left")

        # 创建确认按钮
        button_confirm = tk.Button(self.department_frame, text="确认", command=self.confirm_department_selection)
        button_confirm.pack(side="left")

    def add_department_popup(self):
        department = self.entry_add_department.get()
        if department:
            self.department_list.append(department)
            self.department_listbox_selection.insert(tk.END, department)
            self.update_department_listbox()
            self.entry_add_department.delete(0, tk.END)

    def select_project(self):
        # 创建项目编号选择框
        self.project_frame = ttk.LabelFrame(self.tab_process, text="选择项目编号")
        self.project_frame.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        # 创建项目编号列表框
        self.project_listbox_selection = tk.Listbox(self.project_frame)
        self.project_listbox_selection.pack(side="left")

        # 将已添加的项目编号显示在列表框中
        for project in self.project_list:
            self.project_listbox_selection.insert(tk.END, project)

        # 创建添加项目编号的文本框和按钮
        self.entry_add_project = tk.Entry(self.project_frame, width=30)
        self.entry_add_project.pack(side="left")
        self.entry_add_project.focus_set()
        self.entry_add_project.bind("<Return>", lambda event: self.add_project_popup())
        self.button_add_project_popup = tk.Button(self.project_frame, text="添加项目编号", command=self.add_project_popup)
        self.button_add_project_popup.pack(side="left")

        # 创建确认按钮
        button_confirm = tk.Button(self.project_frame, text="确认", command=self.confirm_project_selection)
        button_confirm.pack(side="left")

    def add_project_popup(self):
        project = self.entry_add_project.get()
        if project:
            self.project_list.append(project)
            self.project_listbox_selection.insert(tk.END, project)
            self.update_project_listbox()
            self.entry_add_project.delete(0, tk.END)

    def confirm_selection(self):
        # 获取用户选择的相关人并更新相关人输入框的值
        selected_users = [self.user_listbox_selection.get(i) for i in self.user_listbox_selection.curselection()]
        self.entry_related_person.config(state='normal')
        self.entry_related_person.delete(0, tk.END)
        self.entry_related_person.insert(0, ",".join(selected_users))
        self.entry_related_person.config(state='readonly')

        # 销毁用户选择框
        self.user_frame.destroy()

    def confirm_department_selection(self):
        # 获取用户选择的承接部门并更新承接部门输入框的值
        selected_department = self.department_listbox_selection.get(tk.ACTIVE)
        self.department_name_var.set(selected_department)

        # 销毁承接部门选择框
        self.department_frame.destroy()

    def confirm_project_selection(self):
        # 获取用户选择的项目编号并更新项目编号输入框的值
        selected_project = self.project_listbox_selection.get(tk.ACTIVE)
        self.project_number_var.set(selected_project)

        # 销毁项目编号选择框
        self.project_frame.destroy()

    def run(self):
        self.root.mainloop()

# 实例化ExcelProcessor并运行程序
processor = ExcelProcessor()
processor.run()
