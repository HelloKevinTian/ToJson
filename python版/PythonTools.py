# -*- coding: utf-8 -*-
import platform, ctypes

STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE= -11
STD_ERROR_HANDLE = -12
FOREGROUND_BLACK = 0x0
FOREGROUND_BLUE = 0x01 # text color contains blue.
FOREGROUND_GREEN= 0x02 # text color contains green.
FOREGROUND_RED = 0x04 # text color contains red.
FOREGROUND_INTENSITY = 0x08 # text color is intensified.
BACKGROUND_BLUE = 0x10 # background color contains blue.
BACKGROUND_GREEN= 0x20 # background color contains green.
BACKGROUND_RED = 0x40 # background color contains red.
BACKGROUND_INTENSITY = 0x80 # background color is intensified.

''' See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winprog/winprog/windows_api_reference.asp
for information on Windows APIs. - www.jb51.net'''
def set_cmd_color(color):
    """(color) -> bit
    Example: set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY)
    """
    std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    ctypes.windll.kernel32.SetConsoleTextAttribute(std_out_handle, color)

def reset_color():
    set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)

def print_red(text):
    if(platform.system() == 'Windows'):
        set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | BACKGROUND_RED)
        print unicode(text, 'utf-8').encode('gbk')
        reset_color()
    else:
        print '\033[1;41m %s \033[0;m' % text

def print_green(text):
    if(platform.system() == 'Windows'):
        set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | BACKGROUND_GREEN)
        print unicode(text, 'utf-8').encode('gbk')
        reset_color()
    else:
        print '\033[1;42m %s \033[0;m' %text

def print_blue(text):
    if(platform.system() == 'Windows'):
        set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | BACKGROUND_BLUE)
        print unicode(text, 'utf-8').encode('gbk')
        reset_color()
    else:
        print '\033[1;44m %s \033[0;m' %text