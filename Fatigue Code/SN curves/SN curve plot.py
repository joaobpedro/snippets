import matplotlib.pyplot as plt
import math
import numpy as np


def B1_sn_curve(N):
    log_a1 = 15.117
    m1 = 4.0

    log_a2 = 17.146
    m2 = 5.0
    
    if N < 10**7:
        return 10**(-1/m1 * math.log(N,10) + 1/m1 * log_a1)
    elif N>=10**7:
        return 10**(-1/m2 * math.log(N,10) + 1/m2 * log_a2)


def curve_5C4N4W(N):
    log_a1 = 13.7
    m1 = 3.32
    return 10**(-1/m1 * math.log(N,10) + 1/m1 * log_a1)

# generate the domain
N = np.logspace(0, 11, 1000)
S =[]
for n in N:
    S.append(curve_5C4N4W(n))


#plot
plt.loglog(N, S, "k", lw=1)
plt.grid(which='both')
plt.title("5C4N4W Reference Curve")
plt.xlabel("Number of cycles")
plt.ylabel("Stress range (MPa)")
plt.ylim(10, 100000)
plt.xlim(1e0, 1e11)
plt.show()