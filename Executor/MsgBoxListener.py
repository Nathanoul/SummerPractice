'''
Allows to create listener object to search and close
popup window with given title (title) with given period (interval)
method start starts listening process
method stop end listening process
method GetMessage gives text in the closed popup
'''
from os import path, remove, mkdir, chdir
import win32com.client as win32
import time
from threading import Thread, Event
import win32gui
import win32con


class MsgBoxListener(Thread):
    def __init__(self, title: str, interval: int):
        Thread.__init__(self)
        self._title = title
        self._interval = interval
        self._stopevent = Event()
        self._message = ''

    def stop(self):
        self._stopevent.set()

    @property
    def isrunning(self):
        return not self._stopevent.is_set()

    def run(self):
        while self.isrunning:
            try:
                time.sleep(self._interval)
                self._closemsgbox()
            except Exception as e:
                print(e, flush=True)

    def _closemsgbox(self):
        # find the top window by title
        hwnd = win32gui.FindWindow(None, self._title)
        if not hwnd: return

        # find child button
        h_btn = win32gui.FindWindowEx(hwnd, None, 'Button', None)
        h_msg = win32gui.FindWindowEx(hwnd, 0, 'Static', None)
        if not h_btn: return

        # save text
        self._message = win32gui.GetWindowText(h_msg)

        # click button
        win32gui.PostMessage(h_btn, win32con.WM_LBUTTONDOWN, None, None)
        time.sleep(0.2)
        win32gui.PostMessage(h_btn, win32con.WM_LBUTTONUP, None, None)
        time.sleep(0.2)

    def GetMessage(self):
        msgrows = self._message.splitlines()
        msg = "    " + "\n    ".join(msgrows)
        return msg
