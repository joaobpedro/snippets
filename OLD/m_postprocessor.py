from OrcFxAPI import*
import glob
import numpy as np
import pandas as pd
from multiprocessing import Pool
from datetime import datetime

#------------------------------------------------------------------------------------------------------------------------------
#SUPPORT FUNCTIONS
#------------------------------------------------------------------------------------------------------------------------------

def find_nearest(array, value):
    array = np.asarray(array)
    idx = (np.abs(array - value)).argmin()
    return array[idx]

def max_couples(vect1, vect2):
    list1 = list(vect1)
    list2 = list(vect2)

    max_1 = max(vect1)
    max_1_2 = vect2[list1.index(max_1)]
    max_2 = max(vect2)
    max_2_1 = vect1[list2.index(max_2)]

    return max_1, max_1_2, max_2, max_2_1

def max_couples_range(vect1, vect2):
    list1 = list(vect1.Max)
    list2 = list(vect2.Max)

    max_1 = max(vect1.Max)
    max_1_2 = vect2.Max[list1.index(max_1)]
    max_2 = max(vect2.Max)
    max_2_1 = vect1.Max[list2.index(max_2)]

    return max_1, max_1_2, max_2, max_2_1

def get_max_in_range(vect1, r1, r2 = None):
    length = list(vect1.X)
    value1 = vect1.Max[length.index(find_nearest(length, r1))]
    if r2!=None:
        value2 = max(vect1.Max[length.index(find_nearest(length, r1)):length.index(find_nearest(length, r2))])
    return max(value1, value2)

def get_min_in_range(vect1, r1, r2 = None):
    length = list(vect1.X)
    value1 = vect1.Min[length.index(find_nearest(length, r1))]
    if r2!=None:
        value2 = min(vect1.Min[length.index(find_nearest(length, r1)):length.index(find_nearest(length, r2))])
    return min(value1, value2)

def get_hs_tp_type(name):
    id = name[:4]
    hs = name.split('Hs=')[1][:5]
    tp = name.split('Tp=')[1][:5]
    if float(hs) == 2.7:
        case_type = 'current'
    else:
        case_type = 'wave'
    
    return id, hs, tp, case_type


#------------------------------------------------------------------------------------------------------------------------------
#END OF SUPPORT FUNCTIONS
#------------------------------------------------------------------------------------------------------------------------------


def Main(name):

    #RANGES
    top_range = [0.0, 10.0]
    bo_range  = [1650.0, 2100.0]
    tdp_range = [2100.0, 2300.0]

    print(f'Processing ->  {name}')
    model = Model()
    model.LoadSimulation(name)

    line = model['Line1']
    stiff = model['Stiffener1']


    #TIME HISTORY
    top_tension = line.TimeHistory('Effective tension', None, oeArcLength(0.7))
    top_curv = line.TimeHistory('Curvature', None, oeArcLength(0.7))
    tdp_tension = line.TimeHistory('Effective tension', None, oeTouchdown)
    top_bend_moment = line.TimeHistory('Bend moment', None, oeArcLength(0.7))
    top_shear_force = line.TimeHistory('Shear Force', None, oeArcLength(0.7))

    bs_bend_moment = stiff.TimeHistory('Bend moment', None, oeEndA)
    bs_shear_force = stiff.TimeHistory('Shear Force', None, oeEndA)

    #RANGE GRAPH
    tension = line.RangeGraph('Effective tension')
    curvature = line.RangeGraph('Curvature')
    tension_BO = line.RangeGraph('Effective tension',None, None, arSpecifiedArclengths(1650.0, 2100.0))
    curvature_BO = line.RangeGraph('Curvature',None, None, arSpecifiedArclengths(1650.0, 2100.0))
    norm_curv = line.RangeGraph('Normalised curvature')

    bs_strain = stiff.RangeGraph('Max bending strain')

    #PROCESS

    total_BM = top_bend_moment + bs_bend_moment
    total_shear_force = top_shear_force + bs_shear_force

    max_top_tension, ass_curv, max_top_curv, ass_tension = max_couples(top_tension,top_curv)
    max_top_util = get_max_in_range(norm_curv, top_range[0], top_range[1])
    max_strain = max(bs_strain.Max)
    max_top_bm, ass_sf, max_top_sf, ass_bm = max_couples(total_BM, total_shear_force)
    min_bo_tension = get_min_in_range(tension, bo_range[0], bo_range[1])
    max_bo_util = get_max_in_range(norm_curv, bo_range[0], bo_range[1])
    max_bo_curv = get_max_in_range(curvature, bo_range[0], bo_range[1])
    min_tdp_tension = get_min_in_range(tension, tdp_range[0], tdp_range[1])
    max_tdp_util = get_max_in_range(norm_curv, tdp_range[0], tdp_range[1])
    max_bo_tension, ass_curv_bo, max_bo_curv_2, ass_tension_bo = max_couples_range(tension_BO,curvature_BO)

    #identifiers:
    id, hs, tp, case_type = get_hs_tp_type(name)

    results = [name,id, hs, tp, case_type,max_top_tension,ass_curv,max_top_curv,ass_tension,max_top_util,max_strain,max_top_bm,ass_sf,max_top_sf,ass_bm,min_bo_tension,max_bo_util,max_bo_curv,min_tdp_tension,max_tdp_util,max_bo_tension, ass_curv_bo, max_bo_curv_2, ass_tension_bo]

    print(f'File processed ->  {name}')

    return results



if __name__ == '__main__':
    Results_ALL = pd.DataFrame(columns=['File Name','ID', 'Hs','Tp','Case Type', 'max top tnesion', 'asso curv', 'max top curv', 'asso tension', 'top max utilization', 'max bs strain', ' max top bm', 'asso sf', 'max top sf', 'asso bm', 'min bo tension', 'max bo util', 'max bo curv', 'min tdp tension', 'max tdp utilization','max_bo_tension', 'ass_curv', 'max_bo_curv', 'ass_tension'])
    filenames = tuple(glob.glob('*.sim'),)

    #number of processes limited to 60 in windows, recomended value of 50 
    with Pool(50) as p:
        results = p.map(Main, filenames)

    for lis in results:
        Results_ALL.loc[len(Results_ALL)] = lis  
    Results_ALL.to_csv('results_ALL_'+datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")+'.csv')



