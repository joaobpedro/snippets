import glob
import shutil
import os
from OrcFxAPI import*
from multiprocessing import Pool

def ChangeDumping(name):
    print(f'Processing -> {name}')
    model = Model(name)
    line = model['Line1']

    #increasing dumping
    line.FullStaticsMinDamping = 75.0
    line.FullStaticsMaxDamping = 80.0
    model.general.StaticsMinDamping = 75.0
    model.general.StaticsMaxDamping = 80.0

    model.SaveData(name)
    return 1

if __name__ == '__main__':
    filenames = glob.glob("*.dat")
    with Pool(20) as p:
        p.map(ChangeDumping, filenames)