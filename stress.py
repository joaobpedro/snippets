from OrcFxAPI import*
import numpy as np
from dataclasses import dataclass

@dataclass
class Field:
    """ class to store field properties like depth """
    name:float = "Yellowtail"
    wd:float = 1790
    fpso:str = "One Guayana"
    numb_umbilicals:float = 3.0

@dataclass
class umb:
    """ class to define umbilical properties """
    name:str = "Umb 1 Yellowtail, static part"
    od:float = 0.215
    ea:float = 773.0
    e_mod:float = 200e3

@dataclass
class tube:
    """ class to define the tube properties"""
    name:str = '2.5in tube'
    id:float = 0.0381
    wt:float = 0.004
    od:float = id + 2.0 * wt
    pi:float = 68.9
    po:float = 0.0


### defining the stress class
class stress():
    """ class to calculate stresses """

    def __init__(self, umb:type(umb), tube:type(tube)):
        """ load the tube and umbilical classes to the stress class upon creation """
        self.tube = tube
        self.umb = umb
        
        # self.mbr = mbr
        # self.tens = tens

    def stress_aec(self) -> float:
        """ calculates the end cap stress, not dependent on the tension vs MBR values """
        return ((self.tube.pi-self.tube.po)*(self.tube.od-2.0*self.tube.wt)**2) / (self.tube.od**2 - (self.tube.od - 2*self.tube.wt)**2)

    def stress_at(self, tens) -> float:
        """ calculates the stress due to tensile loads """
        return (self.umb.e_mod * 0.001 * tens/self.umb.ea)

    def stress_bi(self, mbr) -> float:
        """ Calculates the stress due to bending inside the wall of the tube """
        return (self.umb.e_mod * (0.5 * self.tube.id) / mbr)

    def stress_hi(self) -> float:
        """ calculates teh hoop stress inside the wall """
        return (((self.tube.pi - self.tube.po)*(self.tube.od**2.0 + self.tube.id**2.0)) / (self.tube.od**2.0 - self.tube.id**2.0)-self.tube.po)

    def stress_ri(self) -> float:
        """ Calculates the radial stress inside the wall """
        return -self.tube.pi

    def total_axial_stress_inside(self, tens, mbr) -> float:
        """ Caculates the total axial stress inside the wall """
        return self.stress_aec() + self.stress_at(tens) + self.stress_bi(mbr)

    def stress_eqi(self, tens, mbr) -> float:
        """ Calculates the inside equivalent stress """
        return np.sqrt((self.stress_hi()-self.total_axial_stress_inside(tens,mbr))**2.0 + (self.total_axial_stress_inside(tens,mbr)-self.stress_ri())**2.0 + (self.stress_ri() - self.stress_hi())**2.0) / np.sqrt(2.0)

    def stress_bo(self,mbr) -> float:
        """ Calculates the stress due to bending outside the wall of the tube """
        return self.umb.e_mod * (0.5 * self.tube.od) / mbr

    def stress_ho(self) -> float:
        """ calculates teh hoop stress outside the wall """
        return (self.tube.pi - self.tube.po)*((2.0*self.tube.id**2.0) / (self.tube.od**2.0 - self.tube.id**2.0)) - self.tube.po

    def stress_ro(self) -> float:
        """ Calculates the radial stress outside the wall """
        return -self.tube.po

    def total_axial_stress_outside(self, tens, mbr) -> float:
        """ Caculates the total axial stress outside the wall """
        return self.stress_aec() + self.stress_at(tens) + self.stress_bo(mbr)

    def stress_eqo(self, tens, mbr) -> float:
        """ Calculates the outside equivalent stress """
        return np.sqrt((self.stress_ho()-self.total_axial_stress_outside(tens,mbr))**2.0 + (self.total_axial_stress_outside(tens,mbr)-self.stress_ro())**2.0 + (self.stress_ro() - self.stress_ho())**2.0) / np.sqrt(2.0)

    def get_max_stress(self, tens, mbr):
        """ get the max overall stress for the tube """
        return max(self.stress_eqi(tens,mbr), self.stress_eqo(tens,mbr))



if __name__ == '__main__':
    i_file = open('input_data.csv')
    data = i_file.readlines()
    i_file.close()

    mbr=[]
    tens=[]
    for line in data[1:]:
        mbr.append(float(line.split(',')[0]))
        tens.append(float(line.split(',')[1]))
    
    umb1 = umb()
    tube1 = tube()
    stress1 = stress(umb1, tube1)

    ### validation prints -> very professional programming!
    # print(stress1.stress_aec())
    # print(stress1.stress_at(100))
    # print(stress1.stress_bi(12.83))
    # print(stress1.stress_hi())
    # print(stress1.stress_ri())
    # print(stress1.total_axial_stress_inside(100,12.83))
    # print(stress1.stress_eqi(100,12.83))
    # print(stress1.stress_bo(12.83))
    # print(stress1.stress_ho())
    # print(stress1.stress_ro())
    # print(stress1.total_axial_stress_outside(100.0, 12.83))
    # print(stress1.stress_eqo(100.0,12.83))
    # print(stress1.get_max_stress(100.0,12.83))

    max_stress = [stress1.get_max_stress(tens[i], mbr[i]) for i in range(len(mbr))]
    print(max_stress)