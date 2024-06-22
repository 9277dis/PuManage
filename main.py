import os
import platform
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

import pandas as pd
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

driver = webdriver.Edge()

url = "https://pc.pocketuni.net/chooseSchool"

url2 = "https://pocketuni.net/index.php?app=event&mod=School&act=board&cat=nofinish"
missing_uids = []
mis_uids = ["签到", "签退", "删除"]
sids = []


def set_icon(root, icon_name):
    global icon_path
    try:
        # 根据操作系统选择图标文件
        if platform.system() == 'Windows':
            icon_path = os.path.join(os.path.dirname(__file__), icon_name + '.ico')
        else:
            icon_path = os.path.join(os.path.dirname(__file__), icon_name + '.ico')
        icon_image = Image.open(icon_path)
        # 将 PIL 图像转换为 Tkinter 支持的图像格式
        tk_icon = ImageTk.PhotoImage(icon_image)
        # 设置窗口图标
        root.tk.call('wm', 'iconphoto', root._w, tk_icon)
    except FileNotFoundError:
        messagebox.showerror("Error", f"The file {icon_path} was not found.")


def openurl(urls):
    driver.get(urls)
    driver.maximize_window()


def is_have(uid, mid):
    uid = int(uid)
    # mid转为整数
    mid = int(mid)
    print(f"正在检查学号：{uid}")
    input = "/html/body/div[2]/div/div[3]/div[2]/div[3]/div[1]/div[2]/form/input[2]"
    serch = "/html/body/div[2]/div/div[3]/div[2]/div[3]/div[1]/div[2]/form/input[3]"
    driver.find_element(By.XPATH, input).clear()
    driver.find_element(By.XPATH, input).send_keys(uid)
    driver.find_element(By.XPATH, serch).click()
    cs = "member_tr2"
    elements = driver.find_elements(By.CLASS_NAME, cs)
    # 在新线程里执行
    threading.Thread(target=check_id, args=(elements, uid, mid)).start()


def check_id(elements, uid, mid):
    print(mis_uids[mid])
    # 执行对应操作
    if len(elements) == 0:
        missing_uids.append(uid)
    else:
        id_value = elements[0].get_attribute("id")
        if id_value:
            try:
                id_value = driver.find_element(By.ID, id_value)
                links = id_value.find_elements(By.TAG_NAME, 'a')
                if mid == 2:
                    mid = 1
                    mis_uids[mid] = "删除"
                if len(links) > 1 and links[mid].text == mis_uids[mid]:
                    links[mid].click()
                    try:
                        alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
                        alert.accept()
                        time.sleep(3)  # 等待页面更新（可选）
                    except TimeoutException:
                        print("没有弹出框")
                elif len(links) == 0 or links[mid].text != mis_uids[mid]:
                    pass
            except (IndexError, NoSuchElementException):
                print(f"没有找到{mis_uids[mid]}按钮或其他相关元素")
            except Exception as e:
                print(f"在 check_id 中发生错误: {e}")


def is_have_uid(uid):
    uid = int(uid)
    input = "/html/body/div[2]/div/div[3]/div[2]/div[3]/div[1]/div[2]/form/input[2]"
    serch = "/html/body/div[2]/div/div[3]/div[2]/div[3]/div[1]/div[2]/form/input[3]"
    driver.find_element(By.XPATH, input).clear()
    driver.find_element(By.XPATH, input).send_keys(uid)
    driver.find_element(By.XPATH, serch).click()
    time.sleep(1)
    cs = "member_tr2"
    elements = driver.find_elements(By.CLASS_NAME, cs)
    if len(elements) == 0:
        missing_uids.append(uid)


def end_program():
    driver.quit()
    sys.exit()


class MainWindow:
    def __init__(self, root):
        self.next4_btn = None
        self.re_select_btn = None
        self.file_path = ""
        self.next3_btn = None
        self.root = root
        self.root.title("pu助手")
        self.root.geometry("300x350")
        self.label_tip = tk.Label(self.root, text="请先登录。")
        self.label_tip.pack(side=tk.TOP, fill=tk.X, pady=10)
        self.end_btn = None  # 初始化结束按钮
        self.next1_btn = None  # 初始化下一步按钮
        self.next2_btn = None  # 初始化下一步按钮
        self.after_id = None  # 用于保存after调用的ID
        self.file_path2 = ""
        self.check_login()
        style = ttk.Style(self.root)
        style.configure("My.TButton", foreground="blue", background="blue", font=("Arial", 12, "bold"))

    def on_submit(self):
        openurl(url2)
        self.label_tip.config(text="请打开活动成员管理页面。")
        self.next1_btn = ttk.Button(self.root, text="下一步", style="My.TButton", command=self.on_next1)
        self.next1_btn.pack(pady=10)

    def on_next1(self):
        self.next1_btn.forget()
        self.label_tip.config(text="请选择Excel文件。")
        self.select_file()

    def select_file(self):
        self.label_tip.config(text="正在选择文件...")
        threading.Thread(target=self.thread_select_file).start()

    def thread_select_file(self):
        # 在线程中打开文件对话框，避免阻塞GUI
        self.root.after_idle(self.actual_select_file)

    def actual_select_file(self):
        # 必须在主线程中执行，因为Tkinter的widgets不是线程安全的
        # 显示文件名
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            self.file_path2 = self.file_path
            self.root.after(0, self.on_file_selected(self.file_path2), self.file_path2)
        else:
            self.label_tip.config(text="未选择文件。")
            # 显示签到按钮
            self.show_select_mid_buttons()

    def on_file_selected(self, file_path):
        self.label_tip.config(text=f"已选择文件: {file_path.split('/')[-1]}")
        # 显示签到按钮
        self.show_select_mid_buttons()

    def read_excel(self, file_path, mid):
        try:
            # 使用pandas读取Excel文件，保持数据类型不变
            df = pd.read_excel(file_path, engine='openpyxl' if file_path.endswith('.xlsx') else 'xlrd')
            # 假设"学号"列包含的是字符串类型的学号
            uid_list = df["学号"].tolist()  # 这将保持学号作为字符串
            self.process_uids(uid_list, mid)
        except Exception as e:
            self.label_tip.config(text=f"读取文件时出错: {e}")

    def process_uids(self, uid_list, mid):
        self.label_tip.config(text="正在处理学号...")
        if mid == "":
            threading.Thread(target=lambda: self.check_uid(uid_list)).start()
        else:
            threading.Thread(target=lambda: self.thread_process_uids(uid_list, mid)).start()

    def close_thread(self):
        # 关闭线程
        if self.after_id:
            self.root.after_cancel(self.after_id)

    def thread_process_uids(self, uid_list, mid):
        # 在线程中处理学号，避免阻塞GUI
        for uid in uid_list:
            if uid != "":
                is_have(uid, mid)
                time.sleep(0.5)
        self.root.after(0, self.on_uids_processed)

    def on_uids_processed(self):
        self.label_tip.config(text="处理完成。")

    def show_select_mid_buttons(self):
        # 隐藏之前可能存在的按钮
        for btn in [self.re_select_btn, self.next1_btn, self.next1_btn, self.next2_btn,
                    self.next3_btn, self.next4_btn, self.end_btn]:
            if btn and btn.winfo_ismapped():
                btn.forget()
        self.re_select_btn = ttk.Button(self.root, text="重新选择", style="My.TButton", command=self.select_file)
        self.re_select_btn.pack(pady=10)
        self.next1_btn = ttk.Button(self.root, text="签  到", style="My.TButton",
                                    command=lambda: self.read_excel(self.file_path2, mid=0))
        self.next1_btn.pack(pady=10)
        self.next2_btn = ttk.Button(self.root, text="签  退", style="My.TButton",
                                    command=lambda: self.read_excel(self.file_path2, mid=1))
        self.next2_btn.pack(pady=10)
        self.next3_btn = ttk.Button(self.root, text="删  除", style="My.TButton",
                                    command=lambda: self.read_excel(self.file_path2, mid=2))
        self.next3_btn.pack(pady=10)
        self.next4_btn = ttk.Button(self.root, text="查漏补缺", style="My.TButton",
                                    command=lambda: self.read_excel(self.file_path2, mid=""))
        self.next4_btn.pack(pady=10)
        self.end_btn = ttk.Button(self.root, text="结束程序", style="My.TButton", command=end_program)
        self.end_btn.pack(pady=10)

    def check_uid(self, uid_list):
        for uid in uid_list:
            if uid != '':
                is_have_uid(uid)
                time.sleep(0.5)
        self.label_tip.config(text="查询完成。")
        self.next4_btn.forget()
        self.next4_btn = ttk.Button(self.root, text="保存文件", style='My.TButton', command=self.on_save)
        self.next4_btn.pack(pady=10)

    def check_login(self):
        self.after_id = self.root.after(500, self.check_login_status)

    def check_login_status(self):
        current_url = driver.current_url
        current_path = current_url.split('/')[3] if len(current_url.split('/')) > 3 else None  # 防止索引错误
        if current_path:
            if current_path == 'home':
                self.root.after_cancel(self.after_id)
                self.label_tip.config(text="登录成功。")
                self.on_submit()
            else:
                self.after_id = self.root.after(500, self.check_login_status)
        else:
            print("URL格式可能不正确，无法获取路径。")

    def on_save(self):
        data = [{'学号': uid} for uid in missing_uids]
        df = pd.DataFrame(data)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx;*.xls"),
                       ("All files", "*.*")])

        if file_path:
            df.to_excel(file_path, index=False)
            self.label_tip.config(text="保存成功")
        self.next4_btn.forget()
        self.next4_btn = ttk.Button(self.root, text="查漏补缺", style="My.TButton",
                                    command=lambda: self.check_uid())
        self.next4_btn.pack(pady=10)


def main():
    openurl(url)
    root = tk.Tk()
    root.title("pu助手")
    root.geometry("300x350")
    window = MainWindow(root)
    set_icon(root, 'pu')
    root.attributes('-topmost', 1)
    root.mainloop()


if __name__ == "__main__":
    main()
