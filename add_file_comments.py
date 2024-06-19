import os
import win32com.client
import pythoncom


def set_file_comment(file_path, comment):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")

    # 初始化COM库
    pythoncom.CoInitialize()
    print("program start")

    # 创建Shell对象
    shell = win32com.client.Dispatch("Shell.Application")
    folder_path, file_name = os.path.split(file_path)
    folder = shell.NameSpace(folder_path)
    print("17 done")
    if folder is None:
        raise FileNotFoundError(f"The folder {folder_path} could not be found.")

    item = folder.ParseName(file_name)
    print("22 done")

    if item is None:
        raise FileNotFoundError(f"The file {file_name} could not be found in the folder {folder_path}.")

    # 获取所有的详细属性名称并查找备注属性的索引
    comment_index = -1
    for i in range(266):
        property_name = folder.GetDetailsOf(None, i)
        if property_name.lower() == "comments" or property_name == "备注":
            comment_index = i
            break
    print("34 done")

    if comment_index == -1:
        raise Exception("The 'Comments' property index could not be found.")

    # 使用FolderItem的方法设置文件属性
    folder_item = folder.ParseName(file_name)
    if folder_item is None:
        raise Exception(f"Could not retrieve the folder item for {file_name}.")
    print("43 done")

    # 设置备注
    folder_item.ExtendedProperty(f"System.Comment:{comment}")
    print("47 done")

    # 清理COM库
    pythoncom.CoUninitialize()
    print("final done")


# 示例文件路径和备注
file_path = r"D:\backend_develop\add-file-comments\test.txt"
comment = "This is a sample comment."

# 设置文件备注
set_file_comment(file_path, comment)
