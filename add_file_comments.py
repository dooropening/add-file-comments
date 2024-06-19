import os
import win32api
import win32con
import pythoncom
from win32com.shell import shell, shellcon


def set_file_comment(file_path, comment):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")

    # 初始化COM库
    pythoncom.CoInitialize()

    # 获取IShellItem2对象
    item = shell.SHCreateItemFromParsingName(file_path, None, shell.IID_IShellItem2)

    # 获取属性存储对象
    property_store = item.GetPropertyStore(shellcon.GPS_READWRITE, shell.IID_IPropertyStore)

    # 获取评论属性的PROPERTYKEY
    property_key = shell.PKEY_Comment

    # 设置评论属性
    prop_var = pythoncom.Variant(pythoncom.VT_LPWSTR, comment)
    property_store.SetValue(property_key, prop_var)
    property_store.Commit()

    # 清理COM库
    pythoncom.CoUninitialize()

# 示例文件路径和备注
file_path = r"C:\path\to\your\file.txt"
comment = "This is a sample comment."

# 设置文件备注
set_file_comment(file_path, comment)
