import msvcrt 
import os
import sys
import pathlib

path = pathlib.Path(sys.argv[0]).parent / "Héllo_wörld_jap_日本語Result.txt"
if not "Héllo_wörld_jap_日本語Result.txt" in [file.name for file in path.parent.iterdir()]:
    input("The path we get passed from the context menu as the path of this script seems to be faulty (or the WriteComServerLocationResult.txt isn't in the same path as this script).")
with open(path, "w",encoding="utf-8") as file:
    msvcrt.locking(file.fileno(), msvcrt.LK_NBLCK, 1)  # Lock
    for arg in sys.argv[1:]:
        file.write(arg +"\n")		# Unlock passiert automatisch beim Schließen


