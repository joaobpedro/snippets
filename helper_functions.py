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