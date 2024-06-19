import tkinterdnd2
from tkinter import Tk, Label, messagebox, simpledialog
import os
import zipfile
import sys
sys.path.append(r'.\venv\Lib\site-packages')


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

    def drag_enter(self, event):
        event.widget.configure(bg='yellow')

    def drag_leave(self, event):
        event.widget.configure(bg='white')

    def drop(self, event):
        file_path = event.data
        if os.path.isfile(file_path):
            self.process_file(file_path)
        else:
            messagebox.showerror("错误", "只能拖放单个文件！")

    def process_file(self, file_path):
        comment = simpledialog.askstring("输入", "请输入备注信息:")
        if comment:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            note_file = f"{base_name}备注信息.txt"
            self.create_note_file(file_path, note_file, comment)
            self.compress_files(file_path, note_file)

    def create_note_file(self, source_file, note_file, comment):
        with open(note_file, 'w') as f:
            f.write(comment)
        print(f"创建备注文件 {note_file} 成功")

    def compress_files(self, file_a, note_file):
        base_name = os.path.splitext(os.path.basename(file_a))[0]
        zip_file = f"{base_name}.zip"
        try:
            with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipf.write(file_a, os.path.basename(file_a))
                zipf.write(note_file, os.path.basename(note_file))
            print(f"压缩文件 {zip_file} 成功")
            os.remove(file_a)  # 删除原始文件
            os.remove(note_file)  # 删除备注文件
            print(f"文件 {file_a} 和备注文件 {note_file} 已删除")
        except Exception as e:
            print(f"压缩失败: {e}")
            messagebox.showerror("错误", f"压缩失败: {e}")


def main():
    app = DragDropWindow()
    os.chdir(app.cwd)  # 确保所有文件操作在此目录下进行
    app.mainloop()


if __name__ == "__main__":
    main()
