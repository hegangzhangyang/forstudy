#-*-coding:utf-8-*-

def printHello(func):
    def inner(*args):
        print("hello")
        func(*args)
    return inner