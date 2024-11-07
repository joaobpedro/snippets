import glob
import shutil
import os

filenames = glob.glob("*.dat")
static = False
for name in filenames:
    if os.path.isfile(name[:-4]+".sim"):
        print(f'File exists -> {name}')

        #if exist check size for statics, this is empiric
        og_size = os.path.getsize(name)
        if os.path.getsize(name[:-4]+".sim") < 1.25*og_size:
            static = True
    else:
        try:
            shutil.copyfile(name, './ToRun/'+name)
            print(f'File Copied -> {name}')
        except:
            os.mkdir('./ToRun/')
            shutil.copyfile(name, './ToRun/'+name)
            print(f'File Copied -> {name}')

if static:
    print("WARNING: Potential static simulations, please check")
