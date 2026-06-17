import subprocess
import sys

HOST = "cxd@192.168.0.104"
PASSWORD = "cxdxfx12"

def ssh_run(cmd, timeout=30):
    """通过 ssh 执行命令（使用 sshpass 或 Windows 内置 ssh + 密码）"""
    # 尝试使用 sshpass（如果 Mac 有密码登录 SSH 配置）
    full_cmd = f'ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=NUL -p 22 {HOST} "{cmd}"'
    print(f"\n>>> {cmd}")
    # 使用 plink 或 sshpass 或其他方式传入密码
    # 尝试用 PowerShell 的 ssh + echo 管道输入密码（通常无效）
    # 换种方法：使用 plink.exe (PuTTY) 或 sshpass
    r = subprocess.run(
        full_cmd,
        shell=True,
        capture_output=True,
        text=True,
        timeout=timeout,
        input=PASSWORD + "\n",  # 尝试用 stdin 传密码（大部分 ssh 不接受）
    )
    print("STDOUT:", r.stdout[:2000] if r.stdout else "")
    print("STDERR:", r.stderr[:2000] if r.stderr else "")
    print("ExitCode:", r.returncode)
    return r

if __name__ == "__main__":
    # 第一步：探测基本信息
    ssh_run("uname -a")
    ssh_run("echo HOME=$HOME")
    ssh_run("python3 --version")
    ssh_run("pip3 --version")
    ssh_run("which brew || echo 'no brew'")
    ssh_run("ls ~")
