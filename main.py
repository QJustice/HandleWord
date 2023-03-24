import os
import tkinter.filedialog
import tkinter.messagebox
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage, Frame
from docx import Document


class Application(Frame):
    """gui"""

    def __init__(self, master=None):
        super().__init__(master)
        self.res_path = "Image/"
        self.is_ok = None
        self.count_labe = None
        self.entry_5 = None
        self.entry_4 = None
        self.entry_3 = None
        self.entry_2 = None
        self.entry_1 = None
        # self.OUTPUT_PATH = Path(__file__).parent
        # self.ASSETS_PATH = self.OUTPUT_PATH / Path(r"Image")
        self.button_image_2 = None
        self.button_image_1 = None
        self.entry_image_2 = None
        self.entry_image_1 = None
        self.image_image_1 = None
        self.mode_file = tkinter.StringVar()
        self.mode_file.set("输入或者选择模板名称")
        self.winner_name = "输入获奖者姓名\n(若有多个请用分隔符分割)"
        self.count_label_text = tkinter.StringVar()
        self.count_label_text.set("获奖人数")
        self.flag_name = tkinter.StringVar()
        self.flag_name.set("NAME")
        self.flag_split = tkinter.StringVar()
        self.flag_split.set("|")
        self.fold_name = tkinter.StringVar()
        self.fold_name.set(r"BuildFiles")
        self.root = master
        self.root.geometry("1280x720")
        self.root.configure(bg="#FFFFFF")
        self.root.title("主界面")
        self.root.resizable(False, False)
        self.create_widget()

    def relative_to_assets(self, add_path):
        # return self.ASSETS_PATH / Path(path)
        return self.res_path + add_path

    def create_widget(self):
        """创建UI"""
        canvas = Canvas(
            self.root,
            bg="#FFFFFF",
            height=720,
            width=1280,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )
        self.image_image_1 = PhotoImage(
            file=self.relative_to_assets("image_1.png"))
        canvas.create_image(
            640.0,
            360.0,
            image=self.image_image_1
        )

        canvas.place(x=0, y=0)

        self.entry_image_1 = PhotoImage(
            file=self.relative_to_assets("entry_1.png"))
        canvas.create_image(
            780.0,
            113.5,
            image=self.entry_image_1
        )
        self.entry_1 = Entry(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('软体雅黑', 25),
            textvariable=self.mode_file
        )

        self.entry_1.place(
            x=520.0,
            y=81.0,
            width=520.0,
            height=63.0
        )

        canvas.create_text(
            221.0,
            89.0,
            anchor="nw",
            text="模板word",
            fill="#000000",
            font=("Inter", 40 * -1)
        )

        canvas.create_text(
            228.0,
            312.0,
            anchor="nw",
            text="获奖名单",
            fill="#000000",
            font=("Inter", 40 * -1)
        )

        canvas.create_text(
            228.0,
            490.0,
            anchor="nw",
            text="获奖人数",
            fill="#000000",
            font=("Inter", 40 * -1)
        )

        self.entry_image_2 = PhotoImage(
            file=self.relative_to_assets("entry_2.png"))
        canvas.create_image(
            817.5,
            365.5,
            image=self.entry_image_2
        )
        self.entry_2 = Text(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('软体雅黑', 25),
        )
        self.entry_2.insert('insert', self.winner_name)
        self.entry_2.place(
            x=520.0,
            y=202.0,
            width=595.0,
            height=325.0
        )

        self.count_labe = tkinter.Label(self.root,
                                        textvariable=self.count_label_text,  # 设置文本内容
                                        width=10,  # 设置label的宽度：30
                                        height=2,  # 设置label的高度：10
                                        font=('微软雅黑', 20),  # 设置字体：微软雅黑，字号：18
                                        fg='black',  # 设置前景色：白色
                                        bg='white')  # 设置背景色：灰色

        self.count_labe.place(x=228, y=597)

        self.button_image_1 = PhotoImage(
            file=self.relative_to_assets("button_1.png"))
        button_1 = Button(
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.run(),
            relief="flat"
        )
        button_1.place(
            x=645.0,
            y=617.0,
            width=346.0,
            height=53.0
        )

        self.button_image_2 = PhotoImage(
            file=self.relative_to_assets("button_2.png"))
        button_2 = Button(
            image=self.button_image_2,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.find_files(),
            relief="flat"
        )
        button_2.place(
            x=1049.0,
            y=80.0,
            width=66.0,
            height=72.0
        )
        # button_2按钮旁边加个清空按钮
        button_3 = Button(
            text="清空",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.entry_2.delete("1.0", "end"),
            relief="flat",
            font=('软体雅黑', 25),
        )
        button_3.place(
            x=520.0,
            y=617.0,
            width=66.0,
            height=53.0
        )
        button_4 = Button(
            text="关闭",
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.root.destroy(),
            relief="flat",
            font=('软体雅黑', 25),
        )
        button_4.place(
            x=1050.0,
            y=617.0,
            width=66.0,
            height=53.0
        )
        self.entry_3 = Entry(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('软体雅黑', 25),
            textvariable=self.flag_name
        )

        self.entry_3.place(
            x=50.0,
            y=200.0,
            width=120.0,
            height=63.0
        )

        label_3 = tkinter.Label(self.root,
                                text='姓名标志',  # 设置文本内容
                                width=7,  # 设置label的宽度：7
                                height=2,  # 设置label的高度：2
                                font=('微软雅黑', 20),  # 设置字体：微软雅黑，字号：20
                                fg='black',  # 设置前景色：黑色
                                bg="#D9D9D9")  # 设置背景色：灰色
        label_3.place(x=50, y=100)

        self.entry_4 = Entry(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('软体雅黑', 25),
            textvariable=self.flag_split
        )

        self.entry_4.place(
            x=50.0,
            y=400.0,
            width=120.0,
            height=63.0
        )

        label_4 = tkinter.Label(self.root,
                                text='分隔符',  # 设置文本内容
                                width=7,  # 设置label的宽度：7
                                height=2,  # 设置label的高度：2
                                font=('微软雅黑', 20),  # 设置字体：微软雅黑，字号：20
                                fg='black',  # 设置前景色：黑色
                                bg="#D9D9D9")
        label_4.place(x=50, y=300)

        label_5 = tkinter.Label(self.root,
                                text='文件夹名',  # 设置文本内容
                                width=7,  # 设置label的宽度：7
                                height=2,  # 设置label的高度：2
                                font=('微软雅黑', 20),  # 设置字体：微软雅黑，字号：20
                                fg='black',  # 设置前景色：黑色
                                bg="#D9D9D9")
        label_5.place(x=50, y=500)

        self.entry_5 = Entry(
            bd=0,
            bg="#D9D9D9",
            fg="#000716",
            highlightthickness=0,
            font=('软体雅黑', 25),
            textvariable=self.fold_name
        )
        self.entry_5.place(x=50, y=600, width=120, height=63)

    def find_files(self):
        self.mode_file.set(tkinter.filedialog.askopenfilename())

    def run(self):
        self.winner_name = self.entry_2.get('0.0', 'end')
        winner_name_list = self.winner_name.split(self.flag_split.get())
        self.count_label_text.set(str(len(winner_name_list)))
        answer = tkinter.messagebox.askyesno("提示",
                                             "系统已检测到" + str(len(winner_name_list)) + "个请求，是否开始生成文件")
        if not answer:
            return  # 如果点击的是取消按钮，就不执行下面的代码

        self.is_ok = False
        count = 0
        for item in winner_name_list:
            item = item.replace(" ", "")
            item = item.replace("\r", "")
            item = item.replace("\n", "")
            try:
                docStr = Document(self.entry_1.get())
            except:
                tkinter.messagebox.showwarning("提示", "模板文件不存在")
                return
            children = docStr.element.body.iter()
            for child in children:
                # 通过类型判断目录
                if child.tag.endswith('txbx'):
                    for ci in child.iter():
                        if ci.tag.endswith('main}r'):
                            if ci.text == self.flag_name.get():
                                ci.text = item
                                self.is_ok = True
            try:
                if item != "":
                    # 保存文件
                    folder = os.path.exists(self.entry_5.get())
                    if not folder:
                        os.makedirs(self.entry_5.get())
                    path = self.entry_5.get() + "/" + item + ".docx"
                    # path = self.mode_file.get() + "/" + item + ".docx"
                    docStr.save(path)
                    count += 1
            except:
                tkinter.messagebox.showerror("提示", "保存失败")

        if not self.is_ok:
            tkinter.messagebox.showinfo("提示",
                                        "模板文件中没有找到标记\n当前标记为" + self.flag_name.get() + "\n请检查模板文件")
            return
        else:
            tkinter.messagebox.showinfo("提示", "生成成功,保存在" + self.entry_5.get() + "文件夹。共生成"
                                        + str(count) + "个文件,无效空名文件已被忽略")


if __name__ == '__main__':
    window = Tk()
    app = Application(master=window)
    window.mainloop()
