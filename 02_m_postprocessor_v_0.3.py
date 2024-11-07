from OrcFxAPI import*
import glob
import numpy as np
import pandas as pd
from multiprocessing import Pool
from datetime import datetime
import sys, time, random

def license_handler(func):
    """ Decorator to handle license errors. """

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        for i in range(10):
            try:
                return func(*args, **kwargs)
            except DLLError as error:
                if not error.status == stLicensingError:
                    raise error
                time.sleep(i + random.random())
                print(f"Failed attempt {i + 1} of trying to reach a license.")
        return None

    return wrapper

#------------------------------------------------------------------------------------------------------------------------------
#SUPPORT FUNCTIONS
#------------------------------------------------------------------------------------------------------------------------------

def find_nearest(array, value):
    array = np.asarray(array)
    idx = (np.abs(array - value)).argmin()
    return array[idx]

def max_couples(vect1, vect2, mode='single'):
    list1 = list(vect1)
    list2 = list(vect2)

    max_1 = max(vect1)
    max_1_2 = vect2[list1.index(max_1)]
    max_2 = max(vect2)
    max_2_1 = vect1[list2.index(max_2)]

    if mode == 'single':
        return max_1, max_1_2
    else:
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
@license_handler
def CalculatePairs(orca_object, prop1, prop2, start, end):
    ''' 
    object -> is a line object from Orcaflex - line object
    prop1  -> is te first orcaflex variable to be processed - string
    prop2  -> is the second orcaflex variable to be processd - string
    start  -> number defining the start of the range - int or float
    end    -> number defining the end of the range - int or float
    '''

    #extract range graphs
    prop1_range = orca_object.RangeGraph(prop1, None, None, arSpecifiedArclengths(start, end))
    prop2_range = orca_object.RangeGraph(prop2, None, None, arSpecifiedArclengths(start, end))

    #extract max locations
    prop1_X = prop1_range.X[list(prop1_range.Max).index(max(prop1_range.Max))]
    prop2_X = prop2_range.X[list(prop2_range.Max).index(max(prop2_range.Max))]

    #extract timehistory for locations
    prop1_loc1 = orca_object.TimeHistory(prop1, None, oeArcLength(prop1_X))
    prop2_loc1 = orca_object.TimeHistory(prop2, None, oeArcLength(prop1_X))

    prop1_loc2 = orca_object.TimeHistory(prop1, None, oeArcLength(prop2_X))
    prop2_loc2 = orca_object.TimeHistory(prop2, None, oeArcLength(prop2_X))

    #get pairs
    #pair 1
    max_prop1_loc1, max_prop2_loc1 = max_couples(prop1_loc1, prop2_loc1)

    #pair2
    max_prop2_loc2, max_prop1_loc2 = max_couples(prop2_loc2, prop1_loc2)

    return max_prop1_loc1, max_prop2_loc1, max_prop2_loc2, max_prop1_loc2


@license_handler
def Processor(name):

    #RANGES
    top_range = [0.7, 10.0]
    bo_range  = [1650.0, 2000.0]
    tdp_range = [2000.0, 2300.0]

    print(f'Processing ->  {name}')
    model = Model()
    model.LoadSimulation(name)

    line = model['Line1']
    stiff = model['Stiffener1']


    #TIME HISTORY
    # top_tension = line.TimeHistory('Effective tension', None, oeArcLength(0.7))
    # top_curv = line.TimeHistory('Curvature', None, oeArcLength(0.7))

    top_util = line.RangeGraph('Normalised curvature',None,None,arSpecifiedArclengths(top_range[0], top_range[1]))
    
    tdp_tension = line.TimeHistory('Effective tension', None, oeTouchdown)
    tdp_utilization = line.TimeHistory('Normalised curvature', None, oeTouchdown)
    
    #this takes bend moment and shear force at the end of the cone
    top_bend_moment = line.TimeHistory('Bend moment', None, oeArcLength(top_range[0]))
    top_shear_force = line.TimeHistory('Shear Force', None, oeArcLength(top_range[0]))
    bs_bend_moment = stiff.TimeHistory('Bend moment', None, oeEndA)
    bs_shear_force = stiff.TimeHistory('Shear Force', None, oeEndA)

    #RANGE GRAPH BUOYANCY
    tension_bo = line.RangeGraph('Effective tension',None,None,arSpecifiedArclengths(bo_range[0], bo_range[1]))
    curvature_bo = line.RangeGraph('Curvature',None,None,arSpecifiedArclengths(bo_range[0], bo_range[1]))
    norm_curv_bo = line.RangeGraph('Normalised curvature',None,None,arSpecifiedArclengths(bo_range[0], bo_range[1]))

    bs_strain = stiff.RangeGraph('Max bending strain')

    #PROCESS
    total_BM = top_bend_moment + bs_bend_moment
    total_shear_force = top_shear_force + bs_shear_force

    max_top_tension, ass_curv, max_top_curv, ass_tension = CalculatePairs(line, 'Effective Tension', 'Curvature', top_range[0], top_range[1])
    max_top_util = max(top_util.Max)
    max_strain = max(bs_strain.Max)
    max_top_bm, ass_sf, max_top_sf, ass_bm = max_couples(total_BM, total_shear_force,mode='double')
    min_bo_tension = min(tension_bo.Min)
    max_bo_util = max(norm_curv_bo.Max)
    max_bo_curv = max(curvature_bo.Max)
    min_tdp_tension = min(tdp_tension)
    max_tdp_util = max(tdp_utilization)
    max_bo_tension, ass_curv_bo, max_bo_curv_2, ass_tension_bo = CalculatePairs(line, 'Effective Tension', 'Curvature', bo_range[0], bo_range[1])
    
    #identifiers:
    id, hs, tp, case_type = get_hs_tp_type(name)

    results = [name,id, hs, tp, case_type,max_top_tension,ass_curv,max_top_curv,ass_tension,max_top_util,max_strain,max_top_bm,ass_sf,max_top_sf,ass_bm,min_bo_tension,max_bo_util,max_bo_curv,min_tdp_tension,max_tdp_util,max_bo_tension, ass_curv_bo, max_bo_curv_2, ass_tension_bo]

    print(f'File processed ->  {name}')

    return results


if __name__ == '__main__':
    #input number of processes in the system argv
    # if len(sys.argv) != 2:
    #     raise RuntimeError('Usage -> py -3 m_postprocessor.py number_of_cores')
    # else:
    #     try:
    #         number_of_processors = int(sys.argv[1])
    #     except:
    #         raise TypeError('Sys arg need to be an integer')
    #     if number_of_processors > 50:
    #         raise ValueError('Number of processes needs to be lower than 50')
        

    Results_ALL = pd.DataFrame(columns=['File Name','ID', 'Hs','Tp','Case Type', 'max top tnesion', 'asso curv', 'max top curv', 'asso tension', 'top max utilization', 'max bs strain', ' max top bm', 'asso sf', 'max top sf', 'asso bm', 'min bo tension', 'max bo util', 'max bo curv', 'min tdp tension', 'max tdp utilization','max_bo_tension', 'asso_curv', 'max_bo_curv', 'asso_tension'])
    filenames = glob.glob('*.sim')

    #number of processes limited to 60 in Windows 10, recomended value of 50
    number_of_processors = 10
    with Pool(number_of_processors) as p:
        results = p.map(Processor, filenames)

    for lis in results:
        Results_ALL.loc[len(Results_ALL)] = lis  
    Results_ALL.to_csv('results_ALL_'+datetime.now().strftime("%d_%m_%Y-%I_%M_%S_%p")+'.csv')