import ctypes, sys
from ctypes import wintypes
from pathlib import Path
from PIL import Image
sys.path.insert(0,r'E:\codex\excel-tools')
import tkinter as tk
from excel_unmerge_gui import Application, enable_windows_dpi_awareness
out=Path(r'E:\codex\excel-tools\docs\releases\v1.1\evidence'); out.mkdir(parents=True,exist_ok=True)
user=ctypes.WinDLL('user32',use_last_error=True); gdi=ctypes.WinDLL('gdi32',use_last_error=True)
user.GetDC.argtypes=[wintypes.HWND]; user.GetDC.restype=wintypes.HDC
user.ReleaseDC.argtypes=[wintypes.HWND,wintypes.HDC]
user.PrintWindow.argtypes=[wintypes.HWND,wintypes.HDC,wintypes.UINT]; user.PrintWindow.restype=wintypes.BOOL
gdi.CreateCompatibleDC.argtypes=[wintypes.HDC]; gdi.CreateCompatibleDC.restype=wintypes.HDC
gdi.CreateCompatibleBitmap.argtypes=[wintypes.HDC,ctypes.c_int,ctypes.c_int]; gdi.CreateCompatibleBitmap.restype=wintypes.HBITMAP
gdi.SelectObject.argtypes=[wintypes.HDC,wintypes.HANDLE]; gdi.SelectObject.restype=wintypes.HANDLE
gdi.DeleteObject.argtypes=[wintypes.HANDLE]; gdi.DeleteDC.argtypes=[wintypes.HDC]
class HEADER(ctypes.Structure):
    _fields_=[('size',wintypes.DWORD),('width',wintypes.LONG),('height',wintypes.LONG),('planes',wintypes.WORD),('bits',wintypes.WORD),('compression',wintypes.DWORD),('image_size',wintypes.DWORD),('x',wintypes.LONG),('y',wintypes.LONG),('used',wintypes.DWORD),('important',wintypes.DWORD)]
gdi.GetDIBits.argtypes=[wintypes.HDC,wintypes.HBITMAP,wintypes.UINT,wintypes.UINT,ctypes.c_void_p,ctypes.c_void_p,wintypes.UINT]
enable_windows_dpi_awareness(); root=tk.Tk(); root.attributes('-alpha',0); app=Application(root)
try:
    app.files=(r'C:\报表\华东仓.xlsx',r'C:\报表\华南仓.xlsx')
    app.replace_text(app.file_list,'\n'.join(app.files)); app.start_button.configure(state='normal')
    app.file_states={app.files[0]:'success',app.files[1]:'failed'}; app.update_counts()
    app.status.set(app.task_summary()); app.retry_button.configure(state='normal'); app.open_button.configure(state='normal')
    app.append_result('成功  |  '+app.files[0]+'\n已保存：C:\\报表\\华东仓_拆分填充.xlsx','success')
    app.append_result('失败  |  '+app.files[1]+'\n原因：文件或结果目录不可访问。请关闭正在编辑文件的应用，或把原文件复制到可写目录后重试。','failure')
    for name,size in [('ui-default','840x660'),('ui-minimum','680x540')]:
        root.geometry(size); root.update()
        hwnd=root.winfo_id(); width=root.winfo_width(); height=root.winfo_height()
        src=user.GetDC(hwnd); dc=gdi.CreateCompatibleDC(src); bmp=gdi.CreateCompatibleBitmap(src,width,height); previous=gdi.SelectObject(dc,bmp)
        try:
            ok=user.PrintWindow(hwnd,dc,2)
            info=HEADER(ctypes.sizeof(HEADER),width,-height,1,32,0,0,0,0,0,0)
            data=ctypes.create_string_buffer(width*height*4)
            rows=gdi.GetDIBits(dc,bmp,0,height,data,ctypes.byref(info),0)
            Image.frombytes('RGB',(width,height),data.raw,'raw','BGRX').save(out/(name+'.png'))
            print(name,'capture=',ok,'rows=',rows,'size=',width,height)
        finally:
            gdi.SelectObject(dc,previous);gdi.DeleteObject(bmp);gdi.DeleteDC(dc);user.ReleaseDC(hwnd,src)
finally:
    root.destroy()
