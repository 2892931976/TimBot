# coding: utf-8
import os
import random
import re
import logging
import subprocess
import time
import sys
from ctypes import *
import traceback
import SendKeys
import psutil
import win32api
import win32gui
import win32con
from win32com import client
from suepr_code import super_recognize
from util import *

# 主窗体
MAIN_WINDOW = 'main'
# 登录窗体
LOGIN_WINDOW = 'login'
# TIM根目录
TIM_ROOT = os.path.dirname(os.path.abspath(__file__))

logger = logging.getLogger(__name__)
formatter = logging.Formatter('%(asctime)s %(levelname)-8s: %(message)s')
# 文件日志
file_handler = logging.FileHandler("tim.log")
file_handler.setFormatter(formatter)  # 可以通过setFormatter指定输出格式

# 控制台日志
console_handler = logging.StreamHandler(sys.stdout)
console_handler.formatter = formatter  # 也可以直接给formatter赋值

# 为logger添加的日志处理器
logger.addHandler(file_handler)
logger.addHandler(console_handler)

logger.setLevel(logging.INFO)


class DaMo():
    '''大漠DLL的函数封装'''
    def __init__(self):
        # 注册大漠DLL
        user32 = windll.LoadLibrary("user32.dll")
        self.dm = client.Dispatch("dm.dmsoft")

        # 设置路径
        path_ret = self.dm.SetPath(TIM_ROOT)
        # 添加字库
        font_ret = self.dm.SetDict(0, 'font.txt')
        font_ret = self.dm.SetDict(1, 'num.txt')


class TIM():
    def __init__(self):
        """
        TIM实例初始化
        @param path TIM路径
        """
        # 登陆窗口句柄
        self.login_hwnd = 0
        self.login_dm = DaMo().dm
        # 主窗口句柄
        self.hwnd = 0
        self.main_dm = DaMo().dm
        # 好友管理器句柄
        self.friends_hwnd = 0
        self.friends_dm = DaMo().dm
        # 资料卡句柄
        self.introduct_dm = DaMo().dm

        self.send_setting = None
        self.code_setting = None
        self.sent_friends = 0
        self.logger = logger

    def reset_tim(self):
        # 登陆窗口句柄
        self.login_hwnd = 0
        self.login_dm = DaMo().dm
        # 主窗口句柄
        self.hwnd = 0
        self.main_dm = DaMo().dm
        # 好友管理器句柄
        self.friends_hwnd = 0
        self.friends_dm = DaMo().dm
        # 资料卡句柄
        self.introduct_dm = DaMo().dm
        self.sent_friends = 0

        self.kill_tim()

    def get_tim_hwnd(self):
        '''获得TIM窗口的句柄，使用2中方式获得：
        1. 通过窗口标题获得:"TIM"
        2. 通过进程名称"TIM.exe"获得
        '''
        hwnd = win32gui.FindWindow('TXGuiFoundation', 'TIM')
        return hwnd

    def input_account(self, account):
        # 设定鼠标位置到账号框
        self.login_dm.MoveTo(230, 282)
        self.login_dm.LeftClick()
        # 模拟Ctrl-A + Backspace
        win32api.keybd_event(17, 0, 0, 0)
        win32api.keybd_event(65, 0, 0, 0)
        win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
        self.login_dm.KeyPress(8)
        self.login_dm.SendString(self.login_hwnd, account)

    def input_password(self, password):
        # 设定鼠标位置到密码框
        self.login_dm.MoveTo(230, 310)
        self.login_dm.LeftClick()
        SendKeys.SendKeys(password)

    def login(self, account, password):
        """
        登陆TIM
        @param account QQ账号
        @param password QQ密码
        """
        subprocess.Popen(get_tim_path())

        while True:
            time.sleep(1)
            self.login_hwnd = self.get_tim_hwnd()
            if self.login_hwnd != 0:
                self.logger.info("login window not found, keep looking...")
                break

        bind_ret = self.safe_bind_window(self.login_dm, self.login_hwnd, "normal", "normal", "normal", 0)
        if bind_ret == 1:
            # 输入帐号
            self.input_account(account)
            # 输入密码
            self.input_password(password)
            # 回车登陆
            self.login_dm.KeyPress(13)

            return self.get_login_result()
        else:
            self.logger.error("bind login window failed")
            self.kill_tim()
            return 500

    def safe_bind_window(self, dm, hwnd, display, mouse, keypad, mode):
        """
        安全邦定
        :param dm: 传入的大漠实例
        :param hwnd: 句柄
        :param display:
        :param mouse:
        :param keypad:
        :param mode:
        :return:
        # 0: 失败
        # 1: 成功
        """
        if win32gui.IsWindow(hwnd):
            return dm.BindWindow(hwnd, display, mouse, keypad, mode)
        else:
            return 0

    def exit_login_failed(self):
        """
        登陆失败时退出
        :return:
        """
        self.login_dm.MoveTo(442, 90)
        self.login_dm.LeftClick()

    def exit_tim(self):
        """
        发送完毕正常退出
        :return:
        """
        win32gui.BringWindowToTop(self.hwnd)
        self.main_dm.MoveTo(40, 35)
        self.main_dm.LeftClick()
        self.main_dm.MoveTo(83, 479)
        self.main_dm.LeftClick()

    def kill_tim(self):
        """
        杀死TIM进程
        :return:
        """
        try:
            processes = filter(lambda p: p.name() == 'TIM.exe', psutil.process_iter())
            for p in processes:
                p.terminate()
        except psutil.NoSuchProcess:
            self.logger.error(traceback.format_exc())

    def handle_verify_code(self):
        """ 处理验证码登陆问题"""
        self.login_dm.CaptureJpg(121, 202, 254, 258, 'code.jpg', 100)

        verify_code = super_recognize('your-account', 'your-password', 'code.jpg')
        SendKeys.SendKeys(str(verify_code))
        self.login_dm.KeyPress(13)

    def check_current_window(self):
        """
        检查当前登陆状态
        - login window: 495x470
        - main window: 767x614
        """
        # 获得窗口的尺寸进行判断
        hwnd = self.get_tim_hwnd()
        pos = self.get_window_rect(self.login_dm, hwnd)
        if pos is None:
            return None
        width = pos[3] - pos[1]
        if width > 500:
            return MAIN_WINDOW
        else:
            return LOGIN_WINDOW

    def get_window_rect(self, dm, hwnd):
        x1 = x2 = y1 = y2 = 0

        # 检查TIM窗口是否存在
        if hwnd > 0:
            pos = dm.GetWindowRect(hwnd, x1, y1, x2, y2)
            return pos
        else:
            return None

    def get_login_result(self):
        """
        获得登陆结果
        :return:
        """
        count = 0
        while True:
            time.sleep(1)
            current_window = self.check_current_window()
            if current_window == LOGIN_WINDOW:
                login_result = self.check_login_result()
                if login_result in (400, 401, 402, 403, 404):
                    self.exit_login_failed()
                    return login_result
                elif login_result == 201:
                    self.handle_verify_code()
                else:
                    if count > 15:
                        self.kill_tim()
                        return 500
                    else:
                        count += 1
            elif current_window == MAIN_WINDOW:
                self.logger.debug('logging succeed')
                # 登陆成功，满分100：）
                return 100
            else:
                self.logger.error("bind login window failed")
                self.kill_tim()
                return 500

    def check_login_result(self):
        """
        检测登陆失败结果:
        1. 400, 密码错误
        2. 401, 帐号冻结
        3. 402, 保护模式
        4. 403, 登陆超时
        5. 404, 重复登陆
        5. 200, 登录界面
        6. 201, 验证码
        6. 500, 未知情况
        """
        check_rules = [
            {'key': '密码不正确', 'value': 400, 'color': '000000-000000'},
            {'key': '封停', 'value': 401, 'color': '808080'},
            {'key': '保护', 'value': 402, 'color': '000000-000000'},
            {'key': '回收', 'value': 402, 'color': '000000-000000'},
            {'key': '登录超时', 'value': 403, 'color': '000000-000000'},
            {'key': '重复登录', 'value': 404, 'color': '000000-000000'},
            {'key': '验证码', 'value': 201, 'color': '000000-000000'},
            {'key': '注册帐号', 'value': 200, 'color': '2685E3-2685E3'}
        ]

        for rule in check_rules:
            pos = self.get_text_position(self.login_dm, rule['key'], color=rule['color'])
            if pos[0] == 0:
                return rule['value']
        return 500

    def get_friends_hwnd(self):
        """
        获得好友管理器的句柄
        :return:
        """
        friends_hwnd = self.friends_dm.FindWindow('TXGuiFoundation', '好友管理器'.decode('utf-8').encode('gbk'))
        if friends_hwnd > 0:
            self.logger.debug('found friends management window')
            return friends_hwnd
        else:
            self.logger.debug('not found...')
            return None

    def get_introduction_hwnd(self):
        """
        获得资料卡的句柄
        :return:
        """
        # 最多尝试五次
        count = 0
        while True:
            introduction_hwnd = self.introduct_dm.FindWindow('TXGuiFoundation', '的资料'.decode('utf-8').encode('gbk'))
            if introduction_hwnd != 0:
                self.logger.debug('found introduction window')
                return introduction_hwnd
            else:
                self.logger.warning(introduction_hwnd)
                count += 1
                if count > 5:
                    return None
                else:
                    time.sleep(1)
                    continue

    def change_friend_tab(self):
        pos = self.get_window_rect(self.main_dm, self.hwnd)
        if pos[0] == 1:
            self.main_dm.MoveTo((pos[3]-pos[1])/2, 32)
            self.main_dm.LeftClick()
        else:
            self.logger.error(traceback.format_exc())
            raise Exception('Change friend tab failed')

    def get_text_position(self, dm, text, color='000000-000000'):
        '''根据输入的text查找到其在屏幕上的坐标'''
        # 查找字体
        pos_x = 0
        pos_y = 0

        font_ret = dm.UseDict(0)
        if font_ret == 0:
            return None
        else:
            dm_ret = dm.FindStrFast(0, 0,
                                    dm.GetScreenWidth(),
                                    dm.GetScreenHeight(),
                                    text.decode('utf-8').encode('gbk'),
                                    color, 1, pos_x, pos_y)
            return dm_ret

    def get_current_row(self, i):
        if i < 12:
            return 170 + i * 32
        else:
            return 527

    def get_num_text(self, x1, y1, x2, y2, color='000000-000000'):
        '''根据位置查找到数字字符串'''
        # 切换字库到数字，好友个数格式未：n/m
        font_ret = self.friends_dm.UseDict(1)
        if font_ret == 0:
            return None
        else:
            return self.friends_dm.Ocr(x1, y1, x2, y2, color, 1)

    def get_total_friend_from_management(self, x1, y1, x2, y2):
        '''好友管理器中的好友数目,根据OCR查找到的字符串，‘xxxxxxx(25)'''
        num_text = self.get_num_text(x1, y1, x2, y2, color='FEFEFE')
        if num_text is None:
            return 0
        else:
            num = re.findall(r'\d+', num_text)
            return int(num[0]) if len(num) > 0 else 0

    def open_friend_management_window(self, timeout=1):
        # 切换到好友列表
        self.change_friend_tab()
        time.sleep(5 if self.send_setting is None else float(self.send_setting['login_delay']))
        # 打开好友管理器
        round = 3
        while round > 0:
            count = 10
            while count > 0:
                win32gui.BringWindowToTop(self.hwnd)
                self.main_dm.MoveTo(64, 151)
                self.main_dm.RightClick()
                time.sleep(timeout)
                # 向下移动菜单高度和宽度，点击【好友管理器】
                self.main_dm.MoveR(80, 271)
                self.main_dm.LeftClick()
                time.sleep(0.5)
                friends_hwnd = self.get_friends_hwnd()
                if friends_hwnd is not None:
                    return friends_hwnd
                else:
                    time.sleep(2)
                    count = count -1
            round = round - 1
        return None

    def chat_with_group(self, group_name, materials):
        '''根据指定分组名找到成员并发消息'''
        self.hwnd = self.get_tim_hwnd()
        if self.hwnd == 0:
            self.logger.error('Main window not found')
            return None
        else:
            bind_ret = self.safe_bind_window(self.main_dm, self.hwnd, "normal", "normal", "normal", 0)
            if bind_ret == 1:
                # 打开好友管理器
                self.friends_hwnd = self.open_friend_management_window()
                if self.friends_hwnd is None:
                    self.logger.error(traceback.format_exc())
                    return None
                bind_ret = self.safe_bind_window(self.friends_dm, self.friends_hwnd, 'normal', 'normal', 'windows', 0)
                if bind_ret == 0:
                    self.logger.error(traceback.format_exc())
                    return None
                else:
                    # 获得分组位置
                    # 如果是分组是【我的好友】，颜色为FEFEFE
                    if group_name == '我的好友':
                        pos = self.get_text_position(self.friends_dm, group_name, color='FEFEFE')
                    else:
                        pos = self.get_text_position(self.friends_dm, group_name)
                    if pos[0] == 0:
                        # 选中分组
                        self.friends_dm.MoveTo(pos[1], pos[2])
                        self.friends_dm.LeftClick()
                        time.sleep(2)
                        # 获得该分组好友总数
                        num = self.get_total_friend_from_management(pos[1], pos[2] - 10, pos[1] + 165, pos[2] + 20)
                        logger.info('friends count %d' % num)
                        # 点击第一个好友
                        self.friends_dm.MoveTo(290, 170)
                        self.friends_dm.LeftClick()


                        # 从第一个好友开始发送
                        for i in xrange(0, num):
                            # 按[下]键，使得第一个好友获得焦点
                            self.friends_dm.KeyPress(int('0x28', 16))
                            # 获得当前好友的位置，移动到相应位置并双击打开资料卡
                            pos = self.get_current_row(i)
                            self.friends_dm.MoveTo(290, pos)
                            self.friends_dm.LeftDoubleClick()
                            # 获得资料卡窗口句柄
                            introduction_hwnd = self.get_introduction_hwnd()
                            if introduction_hwnd is None:
                                self.logger.error(traceback.format_exc())
                                return None
                            bind_ret = self.safe_bind_window(self.introduct_dm, introduction_hwnd, 'normal', 'normal', 'normal', 0)
                            if bind_ret == 1:
                                # 311, 267 消息按钮
                                self.introduct_dm.MoveTo(311, 267)
                                self.introduct_dm.LeftClick()
                                self.introduct_dm.UnBindWindow()
                                win32gui.PostMessage(introduction_hwnd, win32con.WM_CLOSE, 0, 0)

                                # 绑定对话窗口
                                win32gui.BringWindowToTop(self.hwnd)
                                # 打开窗口后延迟n秒发送消息
                                time.sleep(1 if self.send_setting is None else float(self.send_setting['send_delay']))
                                for m in materials:
                                    if m.startswith('random'):
                                        count = m.replace('random', '')
                                        emoji_string = random_emoji(int(count))
                                        copy_text_to_clipboard(emoji_string)
                                    elif m.startswith('image='):
                                        copy_to_clipboard(m.split('=')[1])
                                    else:
                                        copy_text_to_clipboard(m.decode('utf-8').encode('gbk'))

                                    win32api.PostMessage(self.hwnd, win32con.WM_PASTE)
                                    # 必须有延迟，否则粘贴板功能会出错
                                    time.sleep(1)
                                win32api.keybd_event(18, 0, 0, 0)
                                win32api.keybd_event(83, 0, 0, 0)
                                win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
                                win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)
                                # 发送后延时n秒关闭窗口
                                time.sleep(1 if self.send_setting is None else float(self.send_setting['close_delay']))
                                # 发送好数量增加
                                self.sent_friends += 1

                                # 将好友管理器窗口置顶
                                win32gui.BringWindowToTop(self.friends_hwnd)
                            else:
                                self.logger.error(traceback.format_exc())
                                return None
                        # 退出TIM
                        self.exit_tim()
                    else:
                        self.logger.warning('group not found %s' % group_name.decode('utf-8').encode('gbk'))
                        # 退出TIM
                        self.exit_tim()
            else:
                self.logger.error('bind main windows failed')
                self.kill_tim()

def random_emoji(length):
    return ''.join(random.sample(["/wx", "/pz", "/se", "/fd", "/dy", "/ll",
                                   "/hx", "/bz", "/shui", "/dk", "/gg", "/fn",
                                   "/tp", "/cy", "/jy", "/ng", "/kuk", "/lengh",
                                   "/zk", "/tuu", "/tx", "/ka", "/baiy", "/am",
                                   "/jie", "/kun", "/jk", "/lh", "/hanx", "/db",
                                   "/fendou", "/zhm", "/yiw", "/xu", "/yun",
                                   "/zhem", "/shuai", "/kl", "/qiao", "/zj",
                                   "/ch", "/kb", "/gz", "/qd", "/huaix", "/zhh",
                                   "/yhh", "/hq", "/bs", "/wq", "/kk", "/yx",
                                   "/qq", "/xia", "/kel", "/cd", "/xig", "/pj",
                                   "/lq", "/pp", "/kf", "/fan", "/zt", "/mg",
                                   "/dx", "/sa", "/xin", "/xs", "/dg", "/shd",
                                   "/zhd", "/dao", "/zq", "/pch", "/bb", "/yl",
                                   "/ty", "/lw", "/yb", "/qiang", "/ruo", "/ws",
                                   "/shl", "/bq", "/gy", "/qt", "/cj", "/aini",
                                   "/bu", "/ok", "/aiq", "/fw", "/tiao", "/fad",
                                   "/oh", "/zhq", "/kt", "/ht", "/hsh", "/jd"], length)).replace(' ', '')


if __name__ == '__main__':
    # 初始化TIM
    tim = TIM()
    tim.reset_tim()
    login_status = tim.login('qq', 'password')

    if login_status == 100:
        try:
            tim.chat_with_group('全部好友', ['你好啊','random3'])
        except BaseException as e:
            logger(e)
    else:
        print(login_status)
