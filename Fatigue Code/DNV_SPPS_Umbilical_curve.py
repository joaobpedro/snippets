import fatpack
import numpy as np
import matplotlib.pyplot as plt
import DNV_SN_Curves as DNV

import sys

curve = fatpack.LinearEnduranceCurve(1.0)
curve.Nc = 14.174*10**17
curve.m = 3.32

N = np.logspace(4, 9, 1000)
S = curve.get_stress(N)

plt.figure(dpi=96)
plt.loglog(N, S, 'k')
plt.title("Equinor Reference SN Curve, 5C4N4W")
plt.xlabel("Number of Cycles")
plt.ylabel("Stress range (MPa)")
plt.xlim(10**4, 10**9)
plt.ylim(100,100000)
plt.grid(which='both')
plt.show()