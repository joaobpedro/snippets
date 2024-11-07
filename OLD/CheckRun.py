import glob
import shutil
import os

filenames = glob.glob("*.dat")

for name in filenames:
    if os.path.isfile(name[:-4]+".sim"):
        print(f'File exists -> {name}')
    else:
        shutil.copyfile(name, './ToRun/'+name)
        print(f'File Copied -> {name}')
