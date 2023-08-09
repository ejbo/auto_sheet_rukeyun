import os
import time

def loop_fun(func, second):
    while True:
        func()
        time.sleep(second)

def run_scrapy():
    os.startfile(r'C:\Users\jzl19\Desktop\keruyun\run.exe')


def main():
    loop_fun(run_scrapy, 5)

main()