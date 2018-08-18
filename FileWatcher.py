# -*-coding: utf-8 -*-
import sys, os
import time

import threading, requests, time


class FileWatcher(threading.Thread):
    def __init__(self, filename, callback, encoding='utf-8'):
        threading.Thread.__init__(self)
        self.filename = filename
        self.callback = callback
        self.encoding = encoding
        self._run = False

    def run(self):
        f = open(self.filename, mode='r', encoding=self.encoding)
        self._run = True
        while self._run:
            line = f.readline()
            if not line:
                # time.sleep(0.1)
                pass
            else:
                self.callback(line)

    def stop(self):
        self._run = False


def CallBack(msg):
    print(msg)


if __name__ == "__main__":
    fw = FileWatcher(filename=r'C:\Ztemp\target.txt', callback=CallBack)
    fw.start()
    time.sleep(100)
    fw.stop()