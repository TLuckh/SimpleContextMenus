import msvcrt 
import os
import sys
import pathlib

path = pathlib.Path(sys.argv[0]).parent / "WriteComServerLocationResult.txt"
if not "WriteComServerLocationResult.txt" in [file.name for file in path.parent.iterdir()]:
    input("The path we get passed from the context menu as the path of this script seems to be faulty (or the WriteComServerLocationResult.txt isn't in the same path as this script).")
with open(path, "w") as file:
    msvcrt.locking(file.fileno(), msvcrt.LK_NBLCK, 1)  # Lock
    file.write(sys.argv[0])		# Unlock passiert automatisch beim Schließen
