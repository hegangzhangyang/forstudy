#-*-coding:utf-8-*-

class DataFormatError(Exception):
    def __init__(self,value):
        self.value=value

    def __str__(self):
        return self.value
