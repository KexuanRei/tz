"""Show the requested five-second success popup after dependency verification."""


def main():
    try:
        import win32com.client
        shell = win32com.client.Dispatch('WScript.Shell')
        shell.Popup('依赖安装成功！请关闭旧服务窗口，再双击 start_ui.bat。', 5, '安装成功', 64)
    except Exception:
        # A popup failure does not undo a successful, verified installation.
        print('依赖安装成功。弹窗未能显示，请重新启动 start_ui.bat。')


if __name__ == '__main__': main()
