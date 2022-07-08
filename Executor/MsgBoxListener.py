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
        self._stop_event = Event()
        self._message = ''

    def stop(self):
        self._stop_event.set()

    @property
    def IsRunning(self):
        return not self._stopevent.isset()

    def run(self):
        while self.isrunning:
            try:
                time.sleep(self._interval)
                self._closemsgbox()
            except Exception as e:
                print(e, flush=True)

    def _CloseMsgbox(self):
        # find the top window by title
        hwnd = win32gui.FindWindow(None, self._title)
        if not hwnd: return

        # find child button
        hbtn = win32gui.FindWindowEx(hwnd, None, 'Button', None)
        hmsg = win32gui.FindWindowEx(hwnd, 0, 'Static', None)
        if not hbtn: return

        # save text
        self._message = win32gui.GetWindowText(hmsg)

        # click button
        win32gui.PostMessage(hbtn, win32con.WM_LBUTTONDOWN, None, None)
        time.sleep(0.2)
        win32gui.PostMessage(hbtn, win32con.WM_LBUTTONUP, None, None)
        time.sleep(0.2)

    def GetMessage(self):
        msgrows = self._message.splitlines()
        msg = "    " + "\n    ".join(msgrows)
        return msg
