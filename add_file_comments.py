import os
import zipfile
import send2trash
from tkinter import Tk, Label, messagebox, simpledialog
import tkinterdnd2
import datetime

"""
软件功能：windows10系统中，对文件或文件夹进行压缩，并添加备注。生成的压缩包名称添生成时间，原文件自动移动到回收站。
"""


class FileProcessor:
    def __init__(self, initial_dir):
        self.initial_dir = initial_dir

    def create_note_file(self, item_path, comment):
        # 在文件或文件夹所在目录创建备注文件
        note_file_path = os.path.join(os.path.dirname(item_path), f"{comment}.txt")
        with open(note_file_path, 'w') as f:
            f.write(comment)
        return note_file_path

    def compress_and_clean_up(self, item_path, note_file_path=None):
        """
        压缩目标路径（文件或文件夹），并清理（移动原文件/文件夹到回收站）
        """
        # 标准化路径
        item_path = os.path.normpath(item_path)
        if note_file_path:
            note_file_path = os.path.normpath(note_file_path)
        if os.path.isdir(item_path):
            base_name = os.path.basename(item_path)
        else:
            base_name = os.path.splitext(os.path.basename(item_path))[0]

        now = datetime.datetime.now()
        formatted_date = now.strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(os.path.dirname(item_path), f"{base_name}_{formatted_date}.zip")

        try:
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if os.path.isdir(item_path):
                    # 压缩文件夹
                    for root, dirs, files in os.walk(item_path):
                        for file in files:
                            if file != os.path.basename(note_file_path):
                                zipf.write(os.path.join(root, file),
                                           os.path.relpath(os.path.join(root, file), item_path))
                    # 显式添加备注文件，确保其被包含在压缩包根目录下
                    if note_file_path:
                        zipf.write(note_file_path, os.path.basename(note_file_path))
                else:
                    # 压缩单个文件
                    zipf.write(item_path, os.path.basename(item_path))
                    if note_file_path:
                        zipf.write(note_file_path, os.path.basename(note_file_path))
            print(f"压缩成功，输出为 {output_file}")

            # 尝试将原文件或文件夹及其备注文件移至回收站
            print(note_file_path)
            print(item_path)
            self.safe_send2trash(note_file_path)
            print('note file path process')
            self.safe_send2trash(item_path)
            print('item path process')
        except Exception as e:
            print(f"压缩失败: {e}")
            messagebox.showerror("错误", f"压缩失败: {e}")

    def safe_send2trash(self, path):
        if os.path.exists(path):
            try:
                send2trash.send2trash(path)
                print(f"已将 {path} 移至回收站")
            except Exception as e:
                print(f"移至回收站失败: {e}")


class DragDropWindow(tkinterdnd2.Tk):
    def __init__(self):
        super().__init__()
        self.title('文件/文件夹备注及打包工具')
        self.geometry('300x200')
        self.cwd = os.getcwd()
        self.processor = FileProcessor(self.cwd)

        self.drop_target = Label(self, text="将文件或文件夹拖放到这里")
        self.drop_target.pack(fill='both', expand=True)

        self.drop_target.drop_target_register(tkinterdnd2.DND_FILES)
        self.drop_target.dnd_bind('<<Drop>>', self.drop)
        self.drop_target.dnd_bind('<<DndEnter>>', self.drag_enter)
        self.drop_target.dnd_bind('<<DndLeave>>', self.drag_leave)

    def drag_enter(self, event):
        event.widget.configure(bg='yellow')

    def drag_leave(self, event):
        event.widget.configure(bg='white')

    def drop(self, event):
        path = event.data
        if os.path.exists(path):
            comment = simpledialog.askstring("输入", "请输入备注信息:")
            if comment:
                note_file_path = self.processor.create_note_file(path, comment)
                self.processor.compress_and_clean_up(path, note_file_path)
        else:
            messagebox.showerror("错误", "文件或文件夹不存在！")


if __name__ == "__main__":
    app = DragDropWindow()
    app.mainloop()
