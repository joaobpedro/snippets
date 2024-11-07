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
    line.FullStaticsMinDamping = 50.0
    line.FullStaticsMaxDamping = 500.0
    line.FullStaticsMaxIterations = int(5000)
    model.general.StaticsMinDamping = 50.0
    model.general.StaticsMaxDamping = 500.0

    model.SaveData(name)
    return 1

if __name__ == '__main__':
    filenames = glob.glob("*.dat")
    with Pool(50) as p:
        p.map(ChangeDumping, filenames)