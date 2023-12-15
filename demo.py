import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ctypes
import webbrowser

# 主窗口宽高/px
MAIN_WINDOW_WIDTH = 650
MAIN_WINDOW_HEIGHT = 300

# 金智教务网址
TEMPLATE_URL = "https://jwgl.wvpn.hrbeu.edu.cn/jwapp/sys/emaphome/portal/index.do"


class App:
    def __init__(self, root) -> None:
        self.root = root
        user32 = ctypes.windll.user32
        self.screen_width = user32.GetSystemMetrics(0)
        self.screen_height = user32.GetSystemMetrics(1)
        pad_width = round((self.screen_width-MAIN_WINDOW_WIDTH)/2)
        pad_height = round((self.screen_height-MAIN_WINDOW_HEIGHT)/2)
        self.root.geometry(
            f"{MAIN_WINDOW_WIDTH}x{MAIN_WINDOW_HEIGHT}+{pad_width}+{pad_height}")
        self.root.title("GPA Calculator v1.0")

        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        self.menu_bar.add_command(label="导入成绩", command=self.import_data)
        self.menu_bar.add_command(label="重置", command=self.reset_data)
        self.menu_bar.add_command(label="方式0", command=self.reload_data_normal)
        self.menu_bar.add_command(label="方式1", command=self.reload_data_func1)
        self.menu_bar.add_command(label="关于", command=self.show_about)

        self.orig_data = pd.DataFrame(
            ["课程名", "总成绩", "课程性质", "学分", "绩点", "考试类型"])
        self.orig_data_num = 0  # 课程总数
        self.course_bx_num = 0  # 必修课程数
        self.course_zx_num = 0  # 专选课程数
        self.course_gx_num = 0  # 公选课程数
        self.orig_data_score = 0  # 总学分
        self.course_bx_score = 0  # 必修学分
        self.course_zx_score = 0  # 专选学分
        self.course_gx_score = 0  # 公选学分
        self.gpa_bx = 0  # 必修课绩点
        self.gpa_all = 0  # 全部课程绩点

        self.hint_lable = tk.Label(
            self.root,
            text=f"仅支持金智教务导出的成绩, 务必导出全部成绩, 点击跳转\n"
                 f"网址: {TEMPLATE_URL}",
            font=('Times New Roman', 12),
            justify=tk.LEFT
        )
        self.hint_lable.place(x=20, y=20)
        self.hint_lable.bind('<Button-1>', self.open_url)

        self.score_lable = tk.Label(
            self.root,
            text=f"导入的课程总数: {self.orig_data_num}\n"
                 f"必修课: {self.course_bx_num}\t"
                 f"专选课: {self.course_zx_num}\t"
                 f"公选课: {self.course_gx_num}",
            font=('Times New Roman', 12),
            justify=tk.LEFT)
        self.score_lable.place(x=20, y=90)

        self.credit_lable = tk.Label(
            self.root,
            text=f"课程总学分: {self.orig_data_score}\n"
                 f"必修课: {self.course_bx_score}\t"
                 f"专选课: {self.course_zx_score}\t"
                 f"公选课: {self.course_gx_score}",
            font=('Times New Roman', 12),
            justify=tk.LEFT)
        self.credit_lable.place(x=20, y=160)

        self.gpa_lable = tk.Label(
            self.root,
            text=f"平均绩点:\n"
                 f"必修课绩点: {self.gpa_bx}/4.00\t"
                 f"全部课程绩点: {self.gpa_all}/4.00",
            font=('Times New Roman', 12),
            justify=tk.LEFT)
        self.gpa_lable.place(x=20, y=230)

    def open_url(self, event):
        if messagebox.askokcancel("提示", "确认使用浏览器打开网址吗?"):
            webbrowser.open_new_tab(TEMPLATE_URL)

    def import_data(self) -> None:
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            try:
                data = pd.read_excel(file_path)
                if not self.check_data(data):
                    messagebox.showerror("错误",
                                         "无法解析Excel文件, 请上传金智教务导出成绩\n"
                                         f"{TEMPLATE_URL}")
                    return
                self.orig_data = data
                self.orig_data_num = len(data)
                self.calc_score(self.orig_data)
                self.reload_data_base()
                messagebox.showinfo("成功", "导入成功")
            except pd.errors.ParserError:
                messagebox.showerror("错误",
                                     "无法解析Excel文件, 请上传金智教务导出成绩\n"
                                     f"{TEMPLATE_URL}")

    def reset_data(self) -> None:
        if messagebox.askokcancel("提示", "确认要清空数据吗?"):
            self.orig_data = pd.DataFrame(
                ["课程名", "总成绩", "课程性质", "学分", "绩点", "考试类型"])
            self.orig_data_num = 0
            self.course_bx_num = 0
            self.course_zx_num = 0
            self.course_gx_num = 0
            self.orig_data_score = 0
            self.course_bx_score = 0
            self.course_zx_score = 0
            self.course_gx_score = 0
            self.gpa_bx = 0
            self.gpa_all = 0
            self.reload_data_base()
            self.reload_gpa()
            messagebox.showinfo("成功", "重置成功")

    def show_about(self) -> None:
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        pad_width = round((self.screen_width-350)/2)
        pad_height = round((self.screen_height-120)/2)
        about_window.geometry(f"350x120+{pad_width}+{pad_height}")

        label = tk.Label(
            about_window,
            text="# 关于\n"
                 "- 作者:\tHaidong Xin\n"
                 "- Email:\txhd0728@hrbeu.edu.cn",
            font=("Times New Roman", 12),
            anchor=tk.W,
            justify=tk.LEFT)
        label.pack(pady=20, padx=20)

    def reload_data_base(self) -> None:
        self.score_lable.config(
            text=f"导入的课程总数: {self.orig_data_num}\n"
                 f"必修课: {self.course_bx_num}\t"
                 f"专选课: {self.course_zx_num}\t"
                 f"公选课: {self.course_gx_num}")
        self.credit_lable.config(
            text=f"课程总学分: {self.orig_data_score}\n"
                 f"必修课: {self.course_bx_score}\t"
                 f"专选课: {self.course_zx_score}\t"
                 f"公选课: {self.course_gx_score}"
        )

    def reload_gpa(self) -> None:
        self.gpa_lable.config(
            text=f"平均绩点:\n"
                 f"必修课绩点: {round(self.gpa_bx,3)}/4.00\t"
                 f"全部课程绩点: {round(self.gpa_all,3)}/4.00"
        )

    def reload_data_normal(self) -> None:
        if not self.orig_data_num:
            messagebox.showerror("错误", "未导入成绩")
            return

        gpa_bx_orig = self.clac_gpa_normal(
            self.orig_data[self.orig_data['课程性质'] == '必修'])
        gpa_all_orig = self.clac_gpa_normal(self.orig_data)

        self.gpa_bx = gpa_bx_orig/self.course_bx_score
        self.gpa_all = gpa_all_orig/self.orig_data_score

        self.reload_gpa()

    def reload_data_func1(self) -> None:
        if not self.orig_data_num:
            messagebox.showerror("错误", "未导入成绩")
            return

        gpa_bx_func1 = self.calc_gpa_func1(
            self.orig_data[self.orig_data['课程性质'] == '必修'])
        gpa_all_func1 = self.calc_gpa_func1(self.orig_data)

        self.gpa_bx = gpa_bx_func1/self.course_bx_score
        self.gpa_all = gpa_all_func1/self.orig_data_score

        self.reload_gpa()

    def check_data(self, df) -> bool:
        required_key = ["课程名", "总成绩", "课程性质", "学分", "绩点", "考试类型"]
        for key in required_key:
            if key not in df.columns:
                return False
        return True

    def calc_score(self, df) -> None:

        # 计算必修课学分
        df_bx = df[df['课程性质'] == '必修']
        df_bx_num = len(df_bx)
        df_bx_sum = df_bx['学分'].sum()
        # 计算专选课学分
        df_zx = df[df['课程性质'] == '任选']
        df_zx_num = len(df_zx)
        df_zx_sum = df_zx['学分'].sum()
        # 计算公选课学分
        df_gx = df[df['课程性质'] == '公选']
        df_gx_num = len(df_gx)
        df_gx_sum = df_gx['学分'].sum()

        self.course_bx_num = df_bx_num
        self.course_zx_num = df_zx_num
        self.course_gx_num = df_gx_num

        self.course_bx_score = df_bx_sum
        self.course_zx_score = df_zx_sum
        self.course_gx_score = df_gx_sum
        self.orig_data_score = df_bx_sum+df_zx_sum+df_gx_sum

    def clac_gpa_normal(self, df) -> float:
        df_orig = pd.DataFrame(['总成绩', '学分', '绩点', '考试类型'])
        df_orig = pd.concat([df_orig, df])
        df_orig['sum'] = df_orig.apply(lambda row: row['学分']*row['绩点'], axis=1)
        return df_orig['sum'].sum()

    def convert_grade_to_credit(self, grade, is_ks) -> float:
        if is_ks:
            if grade >= 90:
                return 4.0
            elif grade >= 85:
                return 3.7
            elif grade >= 82:
                return 3.3
            elif grade >= 78:
                return 3.0
            elif grade >= 75:
                return 2.7
            elif grade >= 72:
                return 2.3
            elif grade >= 68:
                return 2.0
            elif grade >= 64:
                return 1.5
            elif grade >= 60:
                return 1.0
            else:
                return 0.0
        else:
            if grade >= 90:
                return 4.0
            elif grade >= 80:
                return 3.7
            elif grade >= 70:
                return 2.7
            elif grade >= 60:
                return 1.5
            else:
                return 0.0

    def calc_gpa_func1(self, df) -> float:
        df_orig = pd.DataFrame(['总成绩', '学分', '绩点', '考试类型'])
        df_orig = pd.concat([df_orig, df])
        df_orig['sum'] = df_orig.apply(
            lambda row: self.convert_grade_to_credit(row['总成绩'], row['考试类型'] == '考试')*row['学分'], axis=1)
        return df_orig['sum'].sum()


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.resizable(0, 0)
    root.mainloop()
