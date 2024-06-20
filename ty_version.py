import tkinterdnd2
from tkinter import Tk, Label, messagebox, simpledialog
import os
import zipfile
import sys


class FileProcessor:
    def __init__(self, initial_dir):
        self.initial_dir = initial_dir

    def create_note_file(self, source_file, comment):
        # 使用完整路径创建备注文件，避免改变当前工作目录和潜在的安全风险
        base_dir = os.path.dirname(source_file)
        note_file = os.path.join(base_dir, f"{comment}.txt")
        with open(note_file, 'w') as f:
            f.write(comment)
        print(f"创建备注文件 {note_file} 成功")

    def compress_files(self, source_file, note_file):
        # 获取源文件的目录作为输出压缩文件的目录
        base_dir = os.path.dirname(source_file)
        output_file = os.path.join(base_dir, f"{os.path.splitext(os.path.basename(source_file))[0]}.zip")

        try:
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipf.write(source_file, os.path.basename(source_file))
                zipf.write(note_file, os.path.basename(note_file))
            print(f"压缩文件 {output_file} 成功")
            # 删除原始文件和备注文件
            os.remove(source_file)
            os.remove(note_file)
            print(f"文件 {source_file} 和备注文件 {note_file} 已删除")
        except zipfile.BadZipFile:
            print(f"压缩失败: 压缩文件格式不正确")
            messagebox.showerror("错误", f"压缩失败: 压缩文件格式不正确")
        except Exception as e:
            print(f"压缩失败: {e}")
            messagebox.showerror("错误", f"压缩失败: {e}")


class DragDropWindow(tkinterdnd2.Tk):
    def __init__(self):
        super().__init__()
        self.title('文件备注工具')
        self.geometry('300x200')
        self.cwd = os.getcwd()  # 记录窗口启动时的工作目录

        self.drop_target = Label(self, text="将文件拖放到这里")
        self.drop_target.pack(fill='both', expand=True)

        self.drop_target.drop_target_register(tkinterdnd2.DND_FILES)
        self.drop_target.dnd_bind('<<Drop>>', self.drop)
        self.drop_target.dnd_bind('<<DndEnter>>', self.drag_enter)
        self.drop_target.dnd_bind('<<DndLeave>>', self.drag_leave)
        self.processor = FileProcessor(self.cwd)

    def drag_enter(self, event):
        event.widget.configure(bg='yellow')

    def drag_leave(self, event):
        event.widget.configure(bg='white')

    def drop(self, event):
        file_path = event.data
        if os.path.isfile(file_path):
            directory = os.path.split(file_path)[0]
            # 保持工作目录不变，使用完整路径进行文件操作
            self.process_file(file_path)
        else:
            messagebox.showerror("错误", "只能拖放单个文件！")

    def process_file(self, file_path):
        comment = simpledialog.askstring("输入", "请输入备注信息:")
        if comment:
            note_file = os.path.join(os.path.dirname(file_path), f"{comment}.txt")
            self.processor.create_note_file(file_path, comment)
            self.processor.compress_files(file_path, note_file)


def main():
    app = DragDropWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
