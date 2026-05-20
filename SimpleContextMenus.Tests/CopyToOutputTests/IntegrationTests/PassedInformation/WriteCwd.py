import os
import sys

with open("WriteComServerLocationResult.txt","w") as file:
    file.write(os.getcwd())

