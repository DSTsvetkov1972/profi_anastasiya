import os, sys
sys.path.append(os.getcwd())

from colorama import Fore, init, Style
from package.logo import logo_colored

init()
print(Style.BRIGHT)
print(logo_colored)
