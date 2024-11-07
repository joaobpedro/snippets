import fatpack
import numpy as np
import matplotlib.pyplot as plt
import DNV_SN_Curves as DNV

import sys

# Assume that `y` is the data series, we generate one here
y = np.random.normal(0., 30., size=10000)

# Extract the stress ranges by rainflow counting
S = fatpack.find_rainflow_ranges(y)

Sc = 90.0
curve = DNV.DNVGL_EnduranceCurve.in_air("B1")
# DNV.plot_curve(curve.name)
fatigue_damage = curve.find_miner_sum(S)
print(fatigue_damage)