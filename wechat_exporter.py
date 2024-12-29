# -*- coding: utf-8 -*-
import sys
import time
import os
from datetime import datetime
from wxauto import WeChat
import win32gui
import win32con
import win32clipboard

# 设置控制台输出编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')

class WeChatExporter:
    def __init__(self):
        """初始化导出器"""
        self.output_dir = "exported_chats"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        self.wx = None

    def connect_wechat(self):
        """连接到微信"""
        print("正在连接微信...")
        
        # 查找微信窗口
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "微信" in title:
                    windows.append((hwnd, title))
                    
        windows = []
        win32gui.EnumWindows(callback, windows)
        
        if not windows:
            raise Exception("未找到微信窗口，请确保微信已登录并且窗口可见")
            
        # 激活微信窗口
        hwnd = windows[0][0]
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(1)
        
        # 连接到微信
        self.wx = WeChat()
        print("成功连接到微信")

    def get_chat_list(self):
        """获取聊天列表"""
        if not self.wx:
            self.connect_wechat()
            
        sessions = self.wx.GetSessionList()
        if not sessions:
            print("未找到任何聊天会话")
            return []
        return sessions

    def export_chat(self, chat_name=None):
        """导出指定聊天的记录"""
        if not self.wx:
            self.connect_wechat()
            
        if chat_name:
            print(f"\n正在切换到与 {chat_name} 的聊天...")
            self.wx.ChatWith(chat_name)
            time.sleep(1)
        
        # 获取当前聊天名称
        current_chat = self.wx.GetCurrentChatName()
        if not current_chat:
            print("未能获取当前聊天名称")
            return
            
        print(f"正在导出与 {current_chat} 的聊天记录...")
        
        # 滚动到顶部
        for _ in range(5):
            self.wx.SendKeys('{HOME}')
            time.sleep(0.5)
        
        # 全选并复制内容
        self.wx.SendKeys('^a')
        time.sleep(0.5)
        self.wx.SendKeys('^c')
        time.sleep(0.5)
        
        # 获取剪贴板内容
        try:
            win32clipboard.OpenClipboard()
            content = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except:
            try:
                win32clipboard.CloseClipboard()
            except:
                pass
            print("获取聊天内容失败")
            return
            
        if not content:
            print("未能获取到聊天内容")
            return
            
        # 保存到文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.output_dir, f"{current_chat}_{timestamp}.txt")
        filename = filename.replace('/', '_').replace('\\', '_')  # 处理文件名中的非法字符
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"聊天记录已保存到：{filename}")
            print(f"内容长度：{len(content)} 字符")
            return True
        except Exception as e:
            print(f"保存文件失败：{e}")
            return False

    def export_all_chats(self):
        """导出所有聊天记录"""
        chat_list = self.get_chat_list()
        if not chat_list:
            return
            
        print(f"\n找到 {len(chat_list)} 个聊天")
        success_count = 0
        
        for chat in chat_list:
            if self.export_chat(chat):
                success_count += 1
            time.sleep(1)  # 等待一下再处理下一个聊天
            
        print(f"\n导出完成，成功导出 {success_count} 个聊天的记录")

def main():
    try:
        print("=== 微信聊天记录导出工具 ===")
        print("\n请确保：")
        print("1. 微信已登录")
        print("2. 微信窗口可见（不能最小化）")
        input("\n准备好后，按回车键继续...")
        
        exporter = WeChatExporter()
        
        while True:
            print("\n请选择操作：")
            print("1. 导出当前聊天记录")
            print("2. 导出指定聊天记录")
            print("3. 导出所有聊天记录")
            print("0. 退出")
            
            choice = input("\n请输入选项：").strip()
            
            if choice == '1':
                exporter.export_chat()
            elif choice == '2':
                chat_list = exporter.get_chat_list()
                if chat_list:
                    print("\n可用的聊天列表：")
                    for i, chat in enumerate(chat_list, 1):
                        print(f"{i}. {chat}")
                    while True:
                        try:
                            idx = int(input("\n请选择要导出的聊天（输入序号）：")) - 1
                            if 0 <= idx < len(chat_list):
                                exporter.export_chat(chat_list[idx])
                                break
                            else:
                                print("无效的选择，请重试")
                        except ValueError:
                            print("请输入有效的数字")
            elif choice == '3':
                exporter.export_all_chats()
            elif choice == '0':
                print("\n感谢使用！")
                break
            else:
                print("\n无效的选择，请重试")
                
    except Exception as e:
        print(f"\n程序出现错误：{e}")
        
    input("\n按回车键退出...")

if __name__ == '__main__':
    main()
