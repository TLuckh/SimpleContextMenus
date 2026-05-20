import os
import sys

with open("WriteComServerLocationResult.txt","w") as file:

    if len(sys.argv)>1:
        for el in sys.argv[1:]:
            file.write(el +"\n")
