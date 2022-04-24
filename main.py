# _*_ coding: utf-8 _*_

import argparse
from pathlib import Path
from datetime import datetime, timedelta
import time
import winsound
import os
import sys
import shutil
import string
import subprocess

# 需要安装 psutil (pip install psutil), 并在项目中设置 (project -> Python Interpreter)
# import psutil

# 需要安装 pywin32, 并在项目中设置 (project -> Python Interpreter)
import win32process
import win32event


def single_instance():
    exe_name = os.path.basename(sys.argv[0])
    # print(exe_name)

    # tasklist|findstr /i "notepad.exe"
    # 未找到则为空
    cmd = 'tasklist|findstr /i \"' + exe_name + '\"'
    # print(cmd)
    with os.popen(cmd) as result:
        i = 1
        for line in result:    # 结果为空时无 line 不执行
            # 特别注意程序本身有2进程? 结果仅2行即1实例, 4行即2实例, 类推
            if i > 2:
                # print(
                # "重复运行...")
                return 1
            else:
                i = i + 1


def get_parser():
    # 处理命令行参数
    # 1.创建一个解析器
    parser = argparse.ArgumentParser(description='根据命令执行后的返回信息触发关机或重启.')
    # 2.添加参数
    parser.add_argument('--exec_name', default='c:\\windows\\system32\\ipconfig.exe',
                        help='首先执行的命令. 默认 c:\\windows\\system32\\ipconfig.exe')
    parser.add_argument('--cycles', default='0',
                        help='循环执行的次数. 默认 0 无限次')
    parser.add_argument('--interval', default='5',
                        help='循环执行的间隔, 单位秒. 默认 5')
    parser.add_argument('--signature', default='ppp',
                        help='特征字符. 例如判断 ipconfig 是否有 VPN 时, 默认 ppp')
    parser.add_argument('--shooting_trigger', default='0', choices=['0', '1'],
                        help='命中特征字符触发. 默认 0 即未命中时触发')
    parser.add_argument('--operation', default='close', choices=['close', 'shutdown', 'restart'],
                        help='触发的操作. 默认 close 仅关闭辅助, shutdown 关闭电脑, restart 重启电脑')
    parser.add_argument('--operation_delay', default='20',
                        help='触发操作的延时, 如20秒后重启电脑. 默认 20')

    dir_error = 0
    if dir_error == 1:
        print('请使用 -h 或 --help 查看参数详解.')
        time.sleep(5)
        # parser.print_help()
        exit(1)

    return parser


def test_administrator_permissions():
    test_dir = 'c:\\windows\\system32\\testdir'
    try:
        os.makedirs(test_dir)
    except Exception as e:
        print('\n\r')

    if Path(test_dir).exists():
        os.rmdir(test_dir)
        return 1
    else:
        return 0


def open_subprocess(procpath, param=""):
    try:
        if Path(procpath).exists():
            commandline = "\"" + procpath + "\" " + param
            result = os.popen(commandline)
            res = result.read()
            print(res)
            return res
    except Exception as e:
        print(e)


def signature_monitor(signature, result):
    for res in result.splitlines():
        if signature in res:
            print('检测字符命中')
            return True
    else:
        print('检测字符未命中')
        return False


def kill_processes(exe_name, exe_wait=5, pid_wait=2):
    # 先kill exe, 再kill 遗留的 pid
    # image_wait 是每个exe结束后等待秒, 可能其下有多个子进程, 酌情设大点
    # pid_wait 是每pid结束后的等待秒, 多实例多 pid 的情况下不要设置过大

    # tasklist.exe /NH /FI "imagename eq notepad.exe"
    cmd = 'tasklist.exe /NH /FI \"imagename eq ' + exe_name + '\"'
    print(cmd)
    with os.popen(cmd) as result:
        for line in result:
            # 依次有5项: 映像名称, PID, 会话名, 会话#, 内存使用. 注意内存使用K前有空格所以split分割为6个数组元素
            temp = [i for i in line.split(' ') if i != '']
            # print(temp)
            # 无记录时返回 "信息: 没有运行的任务匹配指定标准。", 被分割为2个数组元素
            if len(temp) > 2:
                print(line)
                # taskkill /F /T /IM notepad.exe
                cmd = 'taskkill /F /T /IM ' + exe_name
                print(cmd)
                result = os.popen(cmd)
                time.sleep(exe_wait)
                # exe 只需kill一次
                break

    # tasklist.exe /NH /FI "imagename eq notepad.exe"
    cmd = 'tasklist.exe /NH /FI \"imagename eq ' + exe_name + '\"'
    print(cmd)
    with os.popen(cmd) as result:
        for line in result:
            # 依次有5项: 映像名称, PID, 会话名, 会话#, 内存使用. 注意内存使用K前有空格所以split分割为6个数组元素
            temp = [i for i in line.split(' ') if i != '']
            # print(temp)
            # 无记录时返回 "信息: 没有运行的任务匹配指定标准。", 被分割为2个数组元素
            if len(temp) > 2:
                print(line)
                # 第2项是 PID
                # print(temp[1])
                # 终止进程
                cmd = "taskkill.exe /F /PID " + temp[1]
                print(cmd)
                result = os.popen(cmd)
                time.sleep(pid_wait)


# Python创建新进程的几种方式
# 1.父进程阻塞，且可以控制子进程，推荐用subprocess模块，替换了老的的os.system;os.spawn等，且可以传递startinfo等信息给子进程
# 2.父进程销毁，用子进程替换父进程，用os.exe**,如os.execv,os.exel等系列。注意在调用此函数之后，子进程即刻取得父进程的id,原进程之后的函数皆无法运行，且原父进程的资源，如文件等的所有人也变成了新的子进程。如有特殊必要，可在调用此函数前释放资源
# 3.异步启动新进程，父进程在子进程启动后，不阻塞，继续走自己的路。Windows下可用win32api.WinExec及win32api.ShellExec。win32api.WinExec不会有console窗口，不过如果启动的是bat文件，依然会生成console窗口
# 4.异步启动新进程，父进程在子进程启动后，不阻塞，继续走自己的路。在windows下，同3，可以用win32process.CreateProcess() 和 CreateProcessAsUser()，参数也通同系统API下的CreateProcess，比3好的一点是可以穿很多控制参数及信息，比如使得新启动bat文件也隐藏窗口等
# 5.用阻塞的方式创建一个新进程，如os.system,subprocess等，然后通过设置进程ID或销毁父进程的方法把新的子进程变成一个daemon进程，此方法应该用在linux系统环境中，未测试

def open_process(procpath, param=""):
    try:
        if Path(procpath).exists():
            commandline = "\"" + procpath + "\" " + param
            print(commandline)
            # 注意共10个参数, 最末1个没有使用, LPPROCESS_INFORMATION lpProcessInformation // pointer to PROCESS_INFORMATION
            handle = win32process.CreateProcess(None,  # pointer to name of executable module. (use command line)
                                                commandline,  # pointer to command line string
                                                None,  # process security attributes
                                                None,  # thread security attributes
                                                0,  # handle inheritance flag
                                                0,  # creation flags
                                                None,  # pointer to new environment block
                                                str(Path(procpath).parent),  # pointer to current directory name
                                                win32process.STARTUPINFO())  # pointer to STARTUPINFO
            rc = win32event.WaitForSingleObject(handle[0], 10000)
            print(rc)
        else:
            print("异常退出,请确认是否存在 ", procpath)
    except Exception as e:
        print(e)


if __name__ == '__main__':
    if single_instance():
        print("发现重复实例，本实例即将退出")
        time.sleep(10)
        exit()

    # 取命令行参数
    parser = get_parser()
    args = parser.parse_args()

    if test_administrator_permissions() == 0:
        print("[Error] 没有管理员权限，即将退出！")
        print("[Error] 没有管理员权限，即将退出！")
        print("[Error] 没有管理员权限，即将退出！")
        print("[Error] 没有管理员权限，即将退出！")
        print("[Error] 没有管理员权限，即将退出！")
        time.sleep(10)
        exit()

    # 用无限循环保证控制权, 其实 mainplay.exe 启动时会屏蔽所有 cmd, 所以实际无法持续控制
    while True:
        result_subprocess = open_subprocess('C:\\windows\\system32\\ipconfig.exe')
        if signature_monitor(args.signature, result_subprocess) == int(args.shooting_trigger):
            print('触发操作')
            match args.operation:
                case 'close':
                    print('不关机')
                case 'shutdown':
                    commandline = 'C:\\windows\\system32\\shutdown.exe'
                    parameter = '/s /t ' + args.operation_delay
                    open_process(commandline, parameter)
                case 'restart':
                    commandline = 'C:\\windows\\system32\\shutdown.exe'
                    parameter = '/r /t ' + args.operation_delay
                    open_process(commandline, parameter)

            print('\n\r' + "终止 mainplay.exe")
            kill_processes('mainplay.exe', 40, 2)
            print('\n\r' + "终止 LinkControl.exe")
            kill_processes('linkcontrol.exe', 2, 2)
            print('\n\r' + "终止 dnplayer.exe")
            kill_processes('dnplayer.exe', 5, 2)
            print('\n\r' + "终止 bugreport.exe")
            kill_processes('bugreport.exe', 2, 2)
            print('\n\r' + "终止 adb.exe")
            kill_processes('adb.exe', 2, 2)
            time.sleep(int(args.operation_delay))
        else:
            print('不触发操作')
            time.sleep(int(args.interval))
            os.system('cls')
            # subprocess.call('cls', shell=True)
            print("进度侦测中......")


