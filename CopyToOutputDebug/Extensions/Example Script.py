import os
import sys

print(f"The current working directory is:\n{os.getcwd()}")
print(f"The location of the COM Server is:\n{sys.argv[0]}")
if len(sys.argv)>1:
    print("The following files and folders have been selected:")
    for el in sys.argv[1:]:
        print(el)
input("\nPress 'Enter' to close")
