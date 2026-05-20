import msvcrt 
import os
import sys
import pathlib

path = pathlib.Path(sys.argv[0]).parent / "WriteCwdResult.txt"
if not "WriteCwdResult.txt" in [file.name for file in path.parent.iterdir()]:
    input("The path we get passed from the context menu as the path of this script seems to be faulty (or the WriteComServerLocationResult.txt isn't in the same path as this script).")
with open(path, "w") as file:
    msvcrt.locking(file.fileno(), msvcrt.LK_NBLCK, 1)  # Lock
    file.write(os.getcwd())		# Unlock passiert automatisch beim Schließen

