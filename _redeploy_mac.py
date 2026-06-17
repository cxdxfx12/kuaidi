"""在 Mac 上查找并重新部署 SFeeSystem.app"""
import paramiko
import time

HOST = "192.168.0.104"
USER = "cxd"
PASS = "cxdxfx12"

client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
client.connect(HOST, port=22, username=USER, password=PASS, timeout=30)

def run(cmd, wait=5, verbose=True):
    if verbose:
        print(">>> " + cmd)
    try:
        stdin, stdout, stderr = client.exec_command(cmd)
        out = stdout.read().decode("utf-8", errors="replace").strip()
        err = stderr.read().decode("utf-8", errors="replace").strip()
        rc = stdout.channel.recv_exit_status()
        if out and verbose:
            print(out)
        if err and verbose:
            print("[err] " + err[:100])
        if verbose:
            print("[Exit " + str(rc) + "]")
            print()
        return rc, out, err
    except Exception as e:
        print("  错误: " + str(e))
        return -1, "", str(e)

print("=" * 60)
print("  检查 Mac 上的文件位置")
print("=" * 60)
print()

# 1. 查找所有 SFeeSystem.app
print("[1] 搜索整个硬盘上的 SFeeSystem.app:")
run("mdfind 'kMDItemFSName == \"SFeeSystem.app\"' 2>&1 | head -20", wait=10)

# 2. 检查用户目录
print("[2] 用户目录内容:")
run("ls -la ~/")

# 3. 检查桌面（真正的桌面位置
print("[3] 检查桌面（真正的桌面（注意 Mac 桌面可能在不同位置）:")
run("ls -la ~/Desktop/ 2>&1")
run("ls -la $HOME/Desktop/ 2>&1")
run("ls -la ~/. 2>&1 | head -20")

# 4. 检查 excelbest 目录
print("[4] 检查 excelbest 项目目录:")
run("ls -la ~/excelbest/")

# 5. 检查 dist 目录
print("[5] 检查 dist 目录:")
run("ls -la ~/excelbest/dist/")

# 6. 重新部署
print("[6] 确认原始 app 是否存在:")
rc1, out1, _ = run("ls -la ~/excelbest/dist/SFeeSystem.app/ 2>&1", wait=3, verbose=False)

if rc1 == 0:
    print("[OK] 原始 app 存在，重新部署到桌面:")
    run("rm -rf ~/Desktop/SFeeSystem.app 2>&1")
    run("cp -R ~/excelbest/dist/SFeeSystem.app ~/Desktop/SFeeSystem.app 2>&1")
    run("ls -la ~/Desktop/ 2>&1")
    run("du -sh ~/Desktop/SFeeSystem.app 2>&1")
else:
    print("[警告: 原始 app 不存在，检查其他地方")
    # 尝试从其他地方找
    run("find ~/ -name '*.app' -type d 2>/dev/null | head -20")

# 7. 检查 Finder 是否能看到
print("[7] 检查 Finder 看到的桌面:")
run("ls -la ~/Desktop/")
run("open ~/Desktop/ 2>&1")

# 8. 检查 Spotlight 索引
print("[8] 检查应用程序目录:")
run("ls -la /Applications/ 2>&1 | grep -i fee")

# 9. 确认 Mac 用户
print("[9] 把 app 复制到应用程序目录:")
rc2, _, _ = run("ls -la ~/Desktop/SFeeSystem.app/ 2>&1", wait=3, verbose=False)
if rc2 == 0:
    run("cp -R ~/Desktop/SFeeSystem.app /Applications/ 2>&1")
    run("ls -la /Applications/SFeeSystem.app/ 2>&1")

# 10. 给出使用说明
print()
print("=" * 60)
print("  部署完成！")
print("  你可以通过以下方式打开：")
print("  1) 在 Finder 中打开桌面 桌面：双击 SFeeSystem.app")
print("  2) 在 /Applications/SFeeSystem.app")
print("  3) 通过 Spotlight 搜索 \"大圣\" 或 \"SFeeSystem\"")
print("  4) 终端运行：open ~/Desktop/SFeeSystem.app")
print("=" * 60)

client.close()
