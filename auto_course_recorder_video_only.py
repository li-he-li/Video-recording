"""
只录视频版本（无音频）
用于测试或在没有音频设备时使用
"""

# 这个脚本和 auto_course_recorder.py 完全相同
# 只需要修改第233行的音频部分

# 将以下行注释掉或删除：
#     '-f', 'dshow',  # DirectShow（音频）
#     '-i', 'audio=虚拟音频线',  # 捕获系统音频
#     '-c:a', 'aac',  # 音频编码器
#     '-b:a', '128k',  # 音频比特率

# 或者直接复制 auto_course_recorder.py 的 start_recording 方法
# 并修改为以下内容：

import subprocess
import time

def start_recording_video_only(output_file):
    """只录制视频（无音频）"""
    cmd = [
        'ffmpeg',
        '-y',
        '-f', 'gdigrab',
        '-framerate', '30',
        '-i', 'desktop',
        '-c:v', 'libx264',
        '-preset', 'ultrafast',
        '-crf', '23',
        '-pix_fmt', 'yuv420p',
        output_file
    ]

    try:
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        time.sleep(0.5)

        if process.poll() is None:
            print(f"✓ 开始录制（仅视频）: {output_file}")
            return process
        else:
            print("✗ ffmpeg 启动失败")
            return None

    except FileNotFoundError:
        print("✗ 未找到ffmpeg")
        return None

# 使用方法：
# 1. 打开 auto_course_recorder.py
# 2. 找到第221-283行的 start_recording 方法
# 3. 将 cmd 部分替换为上面的只录视频版本
# 4. 注释掉或删除音频相关参数

print(__doc__)
print("\n快速修复：")
print("1. 打开 auto_course_recorder.py")
print("2. 找到第226-240行的 ffmpeg 命令")
print("3. 删除以下4行：")
print("   - '-f', 'dshow',")
print("   - '-i', 'audio=虚拟音频线',")
print("   - '-c:a', 'aac',")
print("   - '-b:a', '128k',")
print("4. 保存文件")
print("5. 重新运行录制脚本")
