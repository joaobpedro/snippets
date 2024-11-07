# -*- coding: utf-8 -*-
"""
Created on Thu Jan 13 15:06:10 2022
Modified on Wed Mar 16 14:47:59 2022 by daniec
Modified on Wed Jul 21 14:45:00 2022 by daniec

@author: irvinm
v.1.2.1  RAO is included in WaveData and Velocity criteria
v.1.3.2  Bimodal Ochi-Hubble Spectrum for function VRA, RUNVRA, Craete Cases  also add Function for Kizomba project
v.1.3.3  Create Load Cases (Current and Wave) for Aker YellowTail
"""

import copy
import glob
import os
import time
import warnings
from asyncio import subprocess
from collections import defaultdict
from timeit import default_timer as timer

import numpy as np
import OrcFxAPI
import pandas as pd
from pathos.multiprocessing import ProcessingPool, ThreadingPool
from pyrsistent import thaw
from tqdm import tqdm


class Screening(object):
    def __init__(
        self,
        Vessel_name="Vessel1",
        orcafile="BaseCaseVessel.dat",
        ExcelFile="test.xlsx",
        Save=False,
        CreateCases=False,
        MetoceanRAO=True,
        Acc=True,
        Spectrum=0,
        **kwargs,
    ):
        """
        Args:
        Vessel= Orcaflex  model vessel name
        orcafile= Orcaflex  file
        ExcelFile: Data frame of the Load Cases
        Save: True or False if True Save orcaflex.dat files
        CreateCases:True or False  if True creates loadCase spreadsheet otherwise the spreadsheet Data should be create
        MetoceanRAO= True or False if True Create load case spreadsheet incliding RAO information
        Acc=True or False if True Use Acceleration criteria for select cases, otherwise velocity
        Spectrum 0 for Unimodal spectrum 1 for Bimodal Ochi-Hubble
        """

        self.Vessel_name = Vessel_name
        self.orcafile = orcafile
        self.ExcelFile = ExcelFile
        self.Save = Save
        self.CreateCases = CreateCases
        self.MetoceanRAO = MetoceanRAO
        self.Acc = Acc
        self.spect=Spectrum

        sheetname2 = "RAO"
        self.data2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname2, header=0)
        self.list_draft = self.data2
        self.model = OrcFxAPI.Model(self.orcafile)

        sheetname3 = "Parameters"
        self.data3 = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0)

        sheetname4 = "WaveData"
        self.data4 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0)

        sheetname5 = "Current"
        self.data5 = pd.read_excel(self.ExcelFile, sheet_name=sheetname5, header=0, index_col=0)

    def LoadCasesVRA(self):
        """Creates load cases using metocean data and Load the cases in the structure  , it is necessary running to perform any
        screening in time or frequency domain"""
        start = timer()
        warnings.simplefilter("ignore")
        if os.path.isdir("01.VRA") == False:
            os.mkdir("01.VRA")

        file = self.ExcelFile

        if self.spect==0:
            ##Create Data Spread sheet
            if self.MetoceanRAO == False:
                if self.CreateCases == True:
                    labels = [
                        "Case",
                        "Hs",
                        "Tp",
                        "WaveDir",
                        "Gamma",
                        "WaveType",
                        "WaveSpectrumParameter",
                        "WaveSeed",
                        "InitialHeading",
                        "ResponseStormDuration",
                        "ResponseOutputPointx[1]",
                        "ResponseOutputPointy[1]",
                        "ResponseOutputPointz[1]",
                        "Offstet[m]",
                        "CurrentName",
                        "Return Period [years]",
                    ]
                    newnivel = [
                        "",
                        "(m)",
                        "(s)",
                        "goes to (ccw from E)",
                        "",
                        "",
                        "",
                        "",
                        "(ccw from E)",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                    ]

                    newnivel = pd.Series(newnivel).to_frame("").T
                    datAux = pd.DataFrame()
                    datAux = datAux.append(newnivel, ignore_index=True)
                    datAux.columns = labels
                    CaseAux = self.data4.iloc[:, 0:13]  # Select wave environmental data
                    offAux = self.data4.iloc[:, 14:16]
                    offAux.dropna(inplace=True)  # Offset data
                    offAux.set_index(offAux.columns[0], inplace=True)

                    def get_label_Current(a, Rper):
                        if a > 720:
                            print("Direction should be <720")
                        if a > 360:
                            a = a - 360
                        if (a > 348.75 and a <= 360) or (a >= 0 and a < 11.25):
                            label_current = "N{}".format(Rper)
                        elif a > 11.25 and a <= 33.75:
                            label_current = "NNE{}".format(Rper)
                        elif a > 33.75 and a <= 56.25:
                            label_current = "NE{}".format(Rper)
                        elif a > 56.25 and a <= 78.75:
                            label_current = "ENE{}".format(Rper)
                        elif a > 78.75 and a <= 101.25:
                            label_current = "E{}".format(Rper)
                        elif a > 101.25 and a <= 123.75:
                            label_current = "ESE{}".format(Rper)
                        elif a > 123.75 and a <= 146.25:
                            label_current = "SE{}".format(Rper)
                        elif a > 146.25 and a <= 168.75:
                            label_current = "SSE{}".format(Rper)
                        elif a > 168.75 and a <= 191.25:
                            label_current = "SSW{}".format(Rper)
                        elif a > 191.25 and a <= 213.75:
                            label_current = "SW{}".format(Rper)
                        elif a > 213.75 and a <= 236.25:
                            label_current = "WSW{}".format(Rper)
                        elif a > 236.25 and a <= 258.75:
                            label_current = "W{}".format(Rper)
                        elif a > 258.75 and a <= 303.25:
                            label_current = "WNW{}".format(Rper)
                        elif a > 303.25 and a <= 326.25:
                            label_current = "NW{}".format(Rper)
                        elif a > 326.25 and a <= 348.75:
                            label_current = "NNW{}".format(Rper)
                        return label_current

                    data3 = self.data3
                    data4 = self.data4
                    s = 1
                    for k in range(data3.iloc[15, 1]):

                        for i, j in enumerate(CaseAux.index):
                            datAux = datAux.append(datAux.iloc[0, :], ignore_index=True)
                            datAux.iloc[s, 0] = s - 1  # Case
                            #datAux.iloc[s, 1] = data4.iloc[i, 1]  # hs
                            datAux.iloc[s, 1] = data4['Hs[m]'][i]  # hs
                            #datAux.iloc[s, 2] = data4.iloc[i, 2]  # tp
                            datAux.iloc[s, 2] = data4['Tp[s]'][i]  # tp
                            #datAux.iloc[s, 3] = data4.iloc[i, 3]  # wave dir
                            datAux.iloc[s, 3] = data4['DirWave[deg]'][i] # wave dir
                            #datAux.iloc[s, 4] = data4.iloc[i, 4]  # gamma
                            datAux.iloc[s, 4] = data4['Gamma'][i]  # gamma
                            datAux.iloc[s, 5] = data3.iloc[3, 1]  # wavetype
                            datAux.iloc[s, 6] = data3.iloc[4, 1]  # waveparameter
                            datAux.iloc[s, 7] = data3.iloc[6, 1]  # wave seed
                            #datAux.iloc[s, 8] = data4.iloc[i, 5]  # Vessel dir
                            datAux.iloc[s, 8] = data4['Heading'][i]  # Vessel dir
                            datAux.iloc[s, 9] = data3.iloc[5, 1]  # Storm Duration
                            datAux.iloc[s, 10] = data3.iloc[17 + k, 0]  # coord x point
                            datAux.iloc[s, 11] = data3.iloc[17 + k, 1]  # coord y point
                            datAux.iloc[s, 12] = data3.iloc[17 + k, 2]  # coord z point
                            datAux.iloc[s, 13] = offAux.loc[
                                data4['ReturnPeriod'][i], offAux.columns[0]
                            ]  # Offset
                            datAux.iloc[s, 14] = get_label_Current(
                                data4['DirWave[deg]'][i], data4['ReturnPeriod'][i]
                            )  # Current
                            #datAux.iloc[s, 15] = data4.iloc[i, 0]  # return period
                            datAux.iloc[s, 15] = data4['ReturnPeriod'][i]  # return period
                            s = s + 1

                    datAux.set_index(datAux.columns[0], inplace=True)
                    # Saving
                    with pd.ExcelWriter(
                        file, engine="openpyxl", mode="a", if_sheet_exists="replace"
                    ) as writer:

                        try:
                            print("Printing Data Cases")
                            # workBook.remove(workBook['VRA_Results'])
                            # workBook.remove(workBook['VRA_Max'])
                        except:
                            print("Printing Data Cases")
                        finally:
                            print("Printing Data Cases")
                            datAux.to_excel(writer, sheet_name="Data")

                    end0 = timer()
                    print(
                        f"Elapsed time - Data : {end0 - start}",
                    )

                sheetname1 = "Data"
                data = pd.read_excel(file, sheet_name=sheetname1, header=0, index_col=0)
                data.dropna(axis=1, how="all", inplace=True)
                self.data = data.iloc[1:]

            ##Create Data Spread sheet
            if self.MetoceanRAO == True:
            
                labels = [
                    "Case",
                    "Hs",
                    "Tp",
                    "WaveDir",
                    "Gamma",
                    "WaveType",
                    "WaveSpectrumParameter",
                    "WaveSeed",
                    "InitialHeading",
                    "ResponseStormDuration",
                    "ResponseOutputPointx[1]",
                    "ResponseOutputPointy[1]",
                    "ResponseOutputPointz[1]",
                    "Offstet[m]",
                    "CurrentName",
                    "Return Period [years]",
                    "RAO",
                    "Draft",
                ]
                newnivel = [
                    "",
                    "(m)",
                    "(s)",
                    "goes to (ccw from E)",
                    "",
                    "",
                    "",
                    "",
                    "(ccw from E)",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                ]

                newnivel = pd.Series(newnivel).to_frame("").T
                datAux = pd.DataFrame()
                datAux = datAux.append(newnivel, ignore_index=True)
                datAux.columns = labels
                CaseAux = self.data4.iloc[:, 0:13]  # Select wave environmental data
                offAux = self.data4.iloc[:, 14:16]
                offAux.dropna(inplace=True)  # Offset data
                offAux.set_index(offAux.columns[0], inplace=True)

                def get_label_Current(a, Rper):
                    if a > 720:
                        print("Direction should be <720")
                    if a > 360:
                        a = a - 360
                    if (a > 348.75 and a <= 360) or (a >= 0 and a < 11.25):
                        label_current = "N{}".format(Rper)
                    elif a > 11.25 and a <= 33.75:
                        label_current = "NNE{}".format(Rper)
                    elif a > 33.75 and a <= 56.25:
                        label_current = "NE{}".format(Rper)
                    elif a > 56.25 and a <= 78.75:
                        label_current = "ENE{}".format(Rper)
                    elif a > 78.75 and a <= 101.25:
                        label_current = "E{}".format(Rper)
                    elif a > 101.25 and a <= 123.75:
                        label_current = "ESE{}".format(Rper)
                    elif a > 123.75 and a <= 146.25:
                        label_current = "SE{}".format(Rper)
                    elif a > 146.25 and a <= 168.75:
                        label_current = "SSE{}".format(Rper)
                    elif a > 168.75 and a <= 191.25:
                        label_current = "SSW{}".format(Rper)
                    elif a > 191.25 and a <= 213.75:
                        label_current = "SW{}".format(Rper)
                    elif a > 213.75 and a <= 236.25:
                        label_current = "WSW{}".format(Rper)
                    elif a > 236.25 and a <= 258.75:
                        label_current = "W{}".format(Rper)
                    elif a > 258.75 and a <= 303.25:
                        label_current = "WNW{}".format(Rper)
                    elif a > 303.25 and a <= 326.25:
                        label_current = "NW{}".format(Rper)
                    elif a > 326.25 and a <= 348.75:
                        label_current = "NNW{}".format(Rper)
                    return label_current

                data3 = self.data3
                data4 = self.data4
                s = 1
                for k in range(data3.iloc[15, 1]):

                    for i, j in enumerate(CaseAux.index):
                        datAux = datAux.append(datAux.iloc[0, :], ignore_index=True)
                        datAux.iloc[s, 0] = s - 1  # Case
                        #datAux.iloc[s, 1] = data4.iloc[i, 1]  # hs
                        datAux.iloc[s, 1] = data4['Hs[m]'][i]  # hs
                        #datAux.iloc[s, 2] = data4.iloc[i, 2]  # tp
                        datAux.iloc[s, 2] = data4['Tp[s]'][i]  # tp
                        #datAux.iloc[s, 3] = data4.iloc[i, 3]  # wave dir
                        datAux.iloc[s, 3] = data4['DirWave[deg]'][i] # wave dir
                        #datAux.iloc[s, 4] = data4.iloc[i, 4]  # gamma
                        datAux.iloc[s, 4] = data4['Gamma'][i]  # gamma
                        datAux.iloc[s, 5] = data3.iloc[3, 1]  # wavetype
                        datAux.iloc[s, 6] = data3.iloc[4, 1]  # waveparameter
                        datAux.iloc[s, 7] = data3.iloc[6, 1]  # wave seed
                        #datAux.iloc[s, 8] = data4.iloc[i, 5]  # Vessel dir
                        datAux.iloc[s, 8] = data4['Heading'][i]  # Vessel dir
                        datAux.iloc[s, 9] = data3.iloc[5, 1]  # Storm Duration
                        datAux.iloc[s, 10] = data3.iloc[17 + k, 0]  # coord x point
                        datAux.iloc[s, 11] = data3.iloc[17 + k, 1]  # coord y point
                        datAux.iloc[s, 12] = data3.iloc[17 + k, 2]  # coord z point
                        datAux.iloc[s, 13] = offAux.loc[data4['ReturnPeriod'][i], offAux.columns[0]]  # Offset
                        datAux.iloc[s, 14] = get_label_Current(
                            data4['DirWave[deg]'][i], data4['ReturnPeriod'][i]
                        )  # Current
                        #datAux.iloc[s, 15] = data4.iloc[i, 0]  # return period.
                        datAux.iloc[s, 15] = data4['ReturnPeriod'][i]  # return period.
                        #datAux.iloc[s, 16] = data4.iloc[i, 6]  # RAO Name
                        datAux.iloc[s, 16] = data4['RAO'][i]  # RAO Name
                        #datAux.iloc[s, 17] = data4.iloc[i, 7]  # Draft
                        datAux.iloc[s, 17] = data4['Draft'][i]  # Draft
                        s = s + 1

                datAux.set_index(datAux.columns[0], inplace=True)
                # Saving
                with pd.ExcelWriter(
                    file, engine="openpyxl", mode="a", if_sheet_exists="replace"
                ) as writer:

                    try:
                        print("Printing Data Cases")
                        # workBook.remove(workBook['VRA_Results'])
                        # workBook.remove(workBook['VRA_Max'])
                    except:
                        print("Printing Data Cases")
                    finally:
                        print("Printing Data Cases")
                        datAux.to_excel(writer, sheet_name="Data")

                end0 = timer()
                print(
                    f"Elapsed time - Data : {end0 - start}",
                )

                sheetname1 = "Data"
                data = pd.read_excel(file, sheet_name=sheetname1, header=0, index_col=0)
                data.dropna(axis=1, how="all", inplace=True)
                self.data = data.iloc[1:]

        if self.spect==1:
            ##Create Data Spread sheet
            if self.MetoceanRAO == False:

                if self.CreateCases == True:
                    labels = [
                        "Case",
                        "Hs",
                        "Tp",
                        "WaveDir",
                        "Gamma",
                        "WaveType",
                        "WaveSpectrumParameter",
                        "WaveSeed",
                        "InitialHeading",
                        "ResponseStormDuration",
                        "ResponseOutputPointx[1]",
                        "ResponseOutputPointy[1]",
                        "ResponseOutputPointz[1]",
                        "Offstet[m]",
                        "CurrentName",
                        "Return Period [years]",
                        'Hs2',
                        'Tp2',
                        'Gamma2',
                    ]
                    newnivel = [
                        "",
                        "(m)",
                        "(s)",
                        "goes to (ccw from E)",
                        "",
                        "",
                        "",
                        "",
                        "(ccw from E)",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        '',
                        '',
                        '',
                    ]

                    newnivel = pd.Series(newnivel).to_frame("").T
                    datAux = pd.DataFrame()
                    datAux = datAux.append(newnivel, ignore_index=True)
                    datAux.columns = labels
                    CaseAux = self.data4.iloc[:, 0:13]  # Select wave environmental data
                    offAux = self.data4.iloc[:, 14:16]
                    offAux.dropna(inplace=True)  # Offset data
                    offAux.set_index(offAux.columns[0], inplace=True)

                    def get_label_Current(a, Rper):
                        if a > 720:
                            print("Direction should be <720")
                        if a > 360:
                            a = a - 360
                        if (a > 348.75 and a <= 360) or (a >= 0 and a < 11.25):
                            label_current = "N{}".format(Rper)
                        elif a > 11.25 and a <= 33.75:
                            label_current = "NNE{}".format(Rper)
                        elif a > 33.75 and a <= 56.25:
                            label_current = "NE{}".format(Rper)
                        elif a > 56.25 and a <= 78.75:
                            label_current = "ENE{}".format(Rper)
                        elif a > 78.75 and a <= 101.25:
                            label_current = "E{}".format(Rper)
                        elif a > 101.25 and a <= 123.75:
                            label_current = "ESE{}".format(Rper)
                        elif a > 123.75 and a <= 146.25:
                            label_current = "SE{}".format(Rper)
                        elif a > 146.25 and a <= 168.75:
                            label_current = "SSE{}".format(Rper)
                        elif a > 168.75 and a <= 191.25:
                            label_current = "SSW{}".format(Rper)
                        elif a > 191.25 and a <= 213.75:
                            label_current = "SW{}".format(Rper)
                        elif a > 213.75 and a <= 236.25:
                            label_current = "WSW{}".format(Rper)
                        elif a > 236.25 and a <= 258.75:
                            label_current = "W{}".format(Rper)
                        elif a > 258.75 and a <= 303.25:
                            label_current = "WNW{}".format(Rper)
                        elif a > 303.25 and a <= 326.25:
                            label_current = "NW{}".format(Rper)
                        elif a > 326.25 and a <= 348.75:
                            label_current = "NNW{}".format(Rper)
                        return label_current

                    data3 = self.data3
                    data4 = self.data4
                    s = 1
                    for k in range(data3.iloc[15, 1]):

                        for i, j in enumerate(CaseAux.index):
                            datAux = datAux.append(datAux.iloc[0, :], ignore_index=True)
                            datAux.iloc[s, 0] = s - 1  # Case
                            #datAux.iloc[s, 1] = data4.iloc[i, 1]  # hs
                            datAux.iloc[s, 1] = data4['Hs[m]'][i]  # hs
                            #datAux.iloc[s, 2] = data4.iloc[i, 2]  # tp
                            datAux.iloc[s, 2] = data4['Tp[s]'][i]  # tp
                            #datAux.iloc[s, 3] = data4.iloc[i, 3]  # wave dir
                            datAux.iloc[s, 3] = data4['DirWave[deg]'][i] # wave dir
                            #datAux.iloc[s, 4] = data4.iloc[i, 4]  # gamma
                            datAux.iloc[s, 4] = data4['Gamma'][i]  # gamma
                            datAux.iloc[s, 5] = data3.iloc[3, 1]  # wavetype
                            datAux.iloc[s, 6] = data3.iloc[4, 1]  # waveparameter
                            datAux.iloc[s, 7] = data3.iloc[6, 1]  # wave seed
                            #datAux.iloc[s, 8] = data4.iloc[i, 5]  # Vessel dir
                            datAux.iloc[s, 8] = data4['Heading'][i]  # Vessel dir
                            datAux.iloc[s, 9] = data3.iloc[5, 1]  # Storm Duration
                            datAux.iloc[s, 10] = data3.iloc[17 + k, 0]  # coord x point
                            datAux.iloc[s, 11] = data3.iloc[17 + k, 1]  # coord y point
                            datAux.iloc[s, 12] = data3.iloc[17 + k, 2]  # coord z point
                            datAux.iloc[s, 13] = offAux.loc[
                                data4['ReturnPeriod'][i], offAux.columns[0]
                            ]  # Offset
                            datAux.iloc[s, 14] = get_label_Current(
                                data4['DirWave[deg]'][i], data4['ReturnPeriod'][i]
                            )  # Current
                            #datAux.iloc[s, 15] = data4.iloc[i, 0]  # return period
                            datAux.iloc[s, 15] = data4['ReturnPeriod'][i]  # return period
                            datAux.iloc[s, 16] = data4['Hs2[m]'][i]  # Hs2
                            datAux.iloc[s, 17] = data4['Tp2[s]'][i]  # Tp2
                            datAux.iloc[s, 18] = data4['Gamma2'][i]  # Gamma2
                            s = s + 1

                    datAux.set_index(datAux.columns[0], inplace=True)
                    # Saving
                    with pd.ExcelWriter(
                        file, engine="openpyxl", mode="a", if_sheet_exists="replace"
                    ) as writer:

                        try:
                            print("Printing Data Cases")
                            # workBook.remove(workBook['VRA_Results'])
                            # workBook.remove(workBook['VRA_Max'])
                        except:
                            print("Printing Data Cases")
                        finally:
                            print("Printing Data Cases")
                            datAux.to_excel(writer, sheet_name="Data")

                    end0 = timer()
                    print(
                        f"Elapsed time - Data : {end0 - start}",
                    )

                sheetname1 = "Data"
                data = pd.read_excel(file, sheet_name=sheetname1, header=0, index_col=0)
                data.dropna(axis=1, how="all", inplace=True)
                self.data = data.iloc[1:]
            ##Create Data Spread sheet
            if self.MetoceanRAO == True:
            
                labels = [
                        "Case",
                        "Hs",
                        "Tp",
                        "WaveDir",
                        "Gamma",
                        "WaveType",
                        "WaveSpectrumParameter",
                        "WaveSeed",
                        "InitialHeading",
                        "ResponseStormDuration",
                        "ResponseOutputPointx[1]",
                        "ResponseOutputPointy[1]",
                        "ResponseOutputPointz[1]",
                        "Offstet[m]",
                        "CurrentName",
                        "Return Period [years]",
                        'RAO',
                        'Draft',
                        'Hs2',
                        'Tp2',
                        'Gamma2',
                    ]
                newnivel = [
                        "",
                        "(m)",
                        "(s)",
                        "goes to (ccw from E)",
                        "",
                        "",
                        "",
                        "",
                        "(ccw from E)",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        '',
                        '',
                        '',
                        '',
                        '',
                    ]    

                newnivel = pd.Series(newnivel).to_frame("").T
                datAux = pd.DataFrame()
                datAux = datAux.append(newnivel, ignore_index=True)
                datAux.columns = labels
                CaseAux = self.data4.iloc[:, 0:13]  # Select wave environmental data
                offAux = self.data4.iloc[:, 14:16]
                offAux.dropna(inplace=True)  # Offset data
                offAux.set_index(offAux.columns[0], inplace=True)

                def get_label_Current(a, Rper):
                    if a > 720:
                        print("Direction should be <720")
                    if a > 360:
                        a = a - 360
                    if (a > 348.75 and a <= 360) or (a >= 0 and a < 11.25):
                        label_current = "N{}".format(Rper)
                    elif a > 11.25 and a <= 33.75:
                        label_current = "NNE{}".format(Rper)
                    elif a > 33.75 and a <= 56.25:
                        label_current = "NE{}".format(Rper)
                    elif a > 56.25 and a <= 78.75:
                        label_current = "ENE{}".format(Rper)
                    elif a > 78.75 and a <= 101.25:
                        label_current = "E{}".format(Rper)
                    elif a > 101.25 and a <= 123.75:
                        label_current = "ESE{}".format(Rper)
                    elif a > 123.75 and a <= 146.25:
                        label_current = "SE{}".format(Rper)
                    elif a > 146.25 and a <= 168.75:
                        label_current = "SSE{}".format(Rper)
                    elif a > 168.75 and a <= 191.25:
                        label_current = "SSW{}".format(Rper)
                    elif a > 191.25 and a <= 213.75:
                        label_current = "SW{}".format(Rper)
                    elif a > 213.75 and a <= 236.25:
                        label_current = "WSW{}".format(Rper)
                    elif a > 236.25 and a <= 258.75:
                        label_current = "W{}".format(Rper)
                    elif a > 258.75 and a <= 303.25:
                        label_current = "WNW{}".format(Rper)
                    elif a > 303.25 and a <= 326.25:
                        label_current = "NW{}".format(Rper)
                    elif a > 326.25 and a <= 348.75:
                        label_current = "NNW{}".format(Rper)
                    return label_current

                data3 = self.data3
                data4 = self.data4
                s = 1
                for k in range(data3.iloc[15, 1]):

                    for i, j in enumerate(CaseAux.index):
                        datAux = datAux.append(datAux.iloc[0, :], ignore_index=True)
                        datAux.iloc[s, 0] = s - 1  # Case
                        #datAux.iloc[s, 1] = data4.iloc[i, 1]  # hs
                        datAux.iloc[s, 1] = data4['Hs[m]'][i]  # hs
                        #datAux.iloc[s, 2] = data4.iloc[i, 2]  # tp
                        datAux.iloc[s, 2] = data4['Tp[s]'][i]  # tp
                        #datAux.iloc[s, 3] = data4.iloc[i, 3]  # wave dir
                        datAux.iloc[s, 3] = data4['DirWave[deg]'][i] # wave dir
                        #datAux.iloc[s, 4] = data4.iloc[i, 4]  # gamma
                        datAux.iloc[s, 4] = data4['Gamma'][i]  # gamma
                        datAux.iloc[s, 5] = data3.iloc[3, 1]  # wavetype
                        datAux.iloc[s, 6] = data3.iloc[4, 1]  # waveparameter
                        datAux.iloc[s, 7] = data3.iloc[6, 1]  # wave seed
                        #datAux.iloc[s, 8] = data4.iloc[i, 5]  # Vessel dir
                        datAux.iloc[s, 8] = data4['Heading'][i]  # Vessel dir
                        datAux.iloc[s, 9] = data3.iloc[5, 1]  # Storm Duration
                        datAux.iloc[s, 10] = data3.iloc[17 + k, 0]  # coord x point
                        datAux.iloc[s, 11] = data3.iloc[17 + k, 1]  # coord y point
                        datAux.iloc[s, 12] = data3.iloc[17 + k, 2]  # coord z point
                        datAux.iloc[s, 13] = offAux.loc[data4['ReturnPeriod'][i], offAux.columns[0]]  # Offset
                        datAux.iloc[s, 14] = get_label_Current(
                            data4['DirWave[deg]'][i], data4['ReturnPeriod'][i]
                        )  # Current
                        #datAux.iloc[s, 15] = data4.iloc[i, 0]  # return period.
                        datAux.iloc[s, 15] = data4['ReturnPeriod'][i]  # return period.
                        #datAux.iloc[s, 16] = data4.iloc[i, 6]  # RAO Name
                        datAux.iloc[s, 16] = data4['RAO'][i]  # RAO Name
                        #datAux.iloc[s, 17] = data4.iloc[i, 7]  # Draft
                        datAux.iloc[s, 17] = data4['Draft'][i]  # Draft

                        datAux.iloc[s, 18] = data4['Hs2[m]'][i]  # Hs2
                        datAux.iloc[s, 19] = data4['Tp2[s]'][i]  # Tp2
                        datAux.iloc[s, 20] = data4['Gamma2'][i]  # Gamma2
                        s = s + 1

                datAux.set_index(datAux.columns[0], inplace=True)
                # Saving
                with pd.ExcelWriter(
                    file, engine="openpyxl", mode="a", if_sheet_exists="replace"
                ) as writer:

                    try:
                        print("Printing Data Cases")
                        # workBook.remove(workBook['VRA_Results'])
                        # workBook.remove(workBook['VRA_Max'])
                    except:
                        print("Printing Data Cases")
                    finally:
                        print("Printing Data Cases")
                        datAux.to_excel(writer, sheet_name="Data")

                end0 = timer()
                print(
                    f"Elapsed time - Data : {end0 - start}",
                )

                sheetname1 = "Data"
                data = pd.read_excel(file, sheet_name=sheetname1, header=0, index_col=0)
                data.dropna(axis=1, how="all", inplace=True)
                self.data = data.iloc[1:]
            
            
        

            
    def RunVRAGA(self):
        """
        Runs the VRA in frequency domain
        Requirement: Run  LoadCasesVRA
        """
        start = timer()
        list_filename = []
        list_Heave_acc = []
        list_roll = []
        list_pitch = []
        list_rot = []
        list_PerReturn = []
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_draft = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_offset = []
        list_current = []
        list_hs2=[]
        list_tp2=[]
        list_gamma2=[]
        casenumber = 0
        data3=self.data3



        if self.spect==0:
            if self.MetoceanRAO == False:

                list_draft = self.list_draft
                data = self.data
                model = self.model
                Vessel_name = self.Vessel_name
            

                for i, df in tqdm(
                    enumerate(list_draft.iloc[:, 0]), desc="Draft-VRA", total=len(list_draft)
                ):
                    for j in tqdm(data.index, desc="Cases-VRA", total=len(data)):
                        # Set Environment conditions
                        environment = model.environment
                        environment.SelectedWave = "Wave1"
                        environment.WaveDirection = data["WaveDir"][j]  # WaveDir
                        environment.WaveHs = data["Hs"][j]  # Hs
                        environment.WaveType = data["WaveType"][j]  # WaveType
                        environment.WaveJONSWAPParameters = data["WaveSpectrumParameter"][
                            j
                        ]  # JonswapPar
                        environment.WaveGamma = data["Gamma"][j]  # Gamma
                        environment.WaveTp = data["Tp"][j]  # Tp
                        environment.WaveSeed = int(data["WaveSeed"][j])  # WaveSeed
                        environment.WaveNumberOfComponents = int(data3.iloc[7, 1]) #NwaveComp
                        vessel = model[Vessel_name]
                        vessel.InitialHeading = data["InitialHeading"][j]  # WaveInitialHeading
                        vessel.Draught = "{}".format(df)
                        vessel.InitialZ = -list_draft[list_draft.columns[1]][i]
                        vessel.ResponseStormDuration = data["ResponseStormDuration"][
                            j
                        ]  # Stormduration
                        vessel.ResponseOutputPointx[0] = data["ResponseOutputPointx[1]"][j]  # Coord X
                        vessel.ResponseOutputPointy[0] = data["ResponseOutputPointy[1]"][j]  # Coord Y
                        vessel.ResponseOutputPointz[0] = data["ResponseOutputPointz[1]"][j]  # Coord Z
                        # Creating file names
                        casenumber += 1
                        filename = "VRA{:04d}_Hs={:05.2f}_Tp={:05.2f}_{}".format(
                            casenumber, environment.WaveHs, environment.WaveTp, df
                        )
                        vessel.SaveSpectralResponseSpreadsheet(r"./01.VRA/{}.xlsx".format(filename))
                        # Savefiles if True
                        if self.Save == True:
                            model.SaveData(r"./01.VRA/{}.dat".format(filename))
                        # OpenSprectralfile

                        xl = pd.ExcelFile(r"./01.VRA/{}.xlsx".format(filename))
                        sheet = xl.sheet_names[1]
                        aux = xl.parse(sheet_name=sheet, header=6)
                        list_filename.append(filename)
                        # Extracting Results
                        if self.Acc == True:
                            list_Heave_acc.append(aux.iloc[16, 4])
                        else:
                            list_Heave_acc.append(aux.iloc[9, 4])
                        list_roll.append(aux.iloc[2, 5])
                        list_pitch.append(aux.iloc[2, 6])
                        list_rot.append((aux.iloc[2, 6] ** 2 + aux.iloc[2, 5] ** 2) ** 0.5)
                        list_hs.append(data["Hs"][j])  # Hs
                        list_tp.append(data["Tp"][j])  # Tp
                        list_dir.append(data["WaveDir"][j])  # Dir
                        list_gamma.append(data["Gamma"][j])  # Gamma
                        list_wavetype.append(data["WaveType"][j])  # Wavetype
                        list_spectralpar.append(data["WaveSpectrumParameter"][j])  # Jonswap Par
                        list_waveseed.append(int(data["WaveSeed"][j]))  # Seed
                        list_heading.append(data["InitialHeading"][j])  # Initialheading
                        list_x.append(data["ResponseOutputPointx[1]"][j])  # X
                        list_y.append(data["ResponseOutputPointy[1]"][j])  # Y
                        list_z.append(data["ResponseOutputPointz[1]"][j])  # Z
                        list_offset.append(data["Offstet[m]"][j])  # Offset
                        list_current.append(data["CurrentName"][j])  # Current
                        list_PerReturn.append(data["Return Period [years]"][j])  # ReturnPeriod
                        list_rao.append(df)
                        list_initialz.append(-list_draft[list_draft.columns[1]][i])
                        xl.close()

            else:

                data = self.data
                model = self.model
                Vessel_name = self.Vessel_name

                for j in tqdm(data.index, desc="Cases-VRA", total=len(data)):
                        # Set Environment conditions
                        environment = model.environment
                        environment.SelectedWave = "Wave1"
                        environment.WaveDirection = data["WaveDir"][j]  # WaveDir
                        environment.WaveHs = data["Hs"][j]  # Hs
                        environment.WaveType = data["WaveType"][j]  # WaveType
                        environment.WaveJONSWAPParameters = data["WaveSpectrumParameter"][j]  # JonswapPar
                        environment.WaveGamma = data["Gamma"][j]  # Gamma
                        environment.WaveTp = data["Tp"][j]  # Tp
                        environment.WaveSeed = int(data["WaveSeed"][j])  # WaveSeed
                        environment.WaveNumberOfComponents = int(data3.iloc[7, 1])
                        vessel = model[Vessel_name]
                        vessel.InitialHeading = data["InitialHeading"][j]  # WaveInitialHeading
                        vessel.Draught = data["RAO"][j]
                        vessel.InitialZ = -data["Draft"][j]
                        vessel.ResponseStormDuration = data["ResponseStormDuration"][j]  # Stormduration
                        vessel.ResponseOutputPointx[0] = data["ResponseOutputPointx[1]"][j]  # Coord X
                        vessel.ResponseOutputPointy[0] = data["ResponseOutputPointy[1]"][j]  # Coord Y
                        vessel.ResponseOutputPointz[0] = data["ResponseOutputPointz[1]"][j]  # Coord Z
                        # Creating file names
                        casenumber += 1
                        filename = "VRA{:04d}_Hs={:05.2f}_Tp={:05.2f}_{}".format(
                            casenumber, environment.WaveHs, environment.WaveTp, vessel.Draught
                        )
                        vessel.SaveSpectralResponseSpreadsheet(r"./01.VRA/{}.xlsx".format(filename))
                        # Savefiles if True
                        if self.Save == True:
                            model.SaveData(r"./01.VRA/{}.dat".format(filename))
                        # OpenSprectralfile
                        xl = pd.ExcelFile(r"./01.VRA/{}.xlsx".format(filename))
                        sheet = xl.sheet_names[1]
                        aux = xl.parse(sheet_name=sheet, header=6)
                        list_filename.append(filename)
                        # Extracting Results
                        if self.Acc == True:
                            list_Heave_acc.append(aux.iloc[16, 4])
                        else:
                            list_Heave_acc.append(aux.iloc[9, 4])
                        list_roll.append(aux.iloc[2, 5])
                        list_pitch.append(aux.iloc[2, 6])
                        list_rot.append((aux.iloc[2, 6] ** 2 + aux.iloc[2, 5] ** 2) ** 0.5)
                        list_hs.append(data["Hs"][j])  # Hs
                        list_tp.append(data["Tp"][j])  # Tp
                        list_dir.append(data["WaveDir"][j])  # Dir
                        list_gamma.append(data["Gamma"][j])  # Gamma
                        list_wavetype.append(data["WaveType"][j])  # Wavetype
                        list_spectralpar.append(data["WaveSpectrumParameter"][j])  # Jonswap Par
                        list_waveseed.append(int(data["WaveSeed"][j]))  # Seed
                        list_heading.append(data["InitialHeading"][j])  # Initialheading
                        list_x.append(data["ResponseOutputPointx[1]"][j])  # X
                        list_y.append(data["ResponseOutputPointy[1]"][j])  # Y
                        list_z.append(data["ResponseOutputPointz[1]"][j])  # Z
                        list_offset.append(data["Offstet[m]"][j])  # Offset
                        list_current.append(data["CurrentName"][j])  # Current
                        list_PerReturn.append(data["Return Period [years]"][j])  # ReturnPeriod
                        list_rao.append(data["RAO"][j])
                        list_initialz.append(-data["Draft"][j])
                        xl.close()

            # printing
            results = pd.DataFrame(index=list_filename)
            if self.Acc == True:
                results["Z-Acceleration Amplitude [m/s2]"] = list_Heave_acc
            else:
                results["Z-Velocity Amplitude [m/s]"] = list_Heave_acc
            results["Roll Amplitude [deg]"] = list_roll
            results["Pitch Amplitude [deg]"] = list_pitch
            results["Rotation [deg]"] = list_rot
            results["Hs[m]"] = list_hs
            results["Tp[s]"] = list_tp
            results["Wave Direction[deg]"] = list_dir
            results["Gamma"] = list_gamma
            results["Wavetype"] = list_wavetype
            results["WaveSpectrumParameter"] = list_spectralpar
            results["WaveSeed"] = list_waveseed
            results["InitialHeading"] = list_heading
            results["x[m]"] = list_x
            results["y[m]"] = list_y
            results["z[m]"] = list_z
            results["Offset[m]"] = list_offset
            results["CurrenName"] = list_current
            results["Return Period"] = list_PerReturn
            results["RAO"] = list_rao
            results["Draft"] = list_initialz

        if self.spect==1:
            if self.MetoceanRAO == False:

                list_draft = self.list_draft
                data = self.data
                model = self.model
                Vessel_name = self.Vessel_name
            

                for i, df in tqdm(
                    enumerate(list_draft.iloc[:, 0]), desc="Draft-VRA", total=len(list_draft)
                ):
                    for j in tqdm(data.index, desc="Cases-VRA", total=len(data)):
                        # Set Environment conditions
                        environment = model.environment
                        environment.SelectedWave = "Wave1"
                        environment.WaveDirection = data["WaveDir"][j]  # WaveDir
                        environment.WaveType = data["WaveType"][j]  # WaveType
                        environment.WaveOchiHubbleParameters = data["WaveSpectrumParameter"][j]  # JonswapPar
                        environment.WaveHs1 = data["Hs"][j]  # Hs1
                        environment.Wavefm1 = 1/data["Tp"][j]  # Tp1
                        environment.WaveLambda1 = data["Gamma"][j]  # Gamma1
                        environment.WaveHs2 = data["Hs2"][j]  # Hs2
                        environment.Wavefm2 = 1/data["Tp2"][j]  # Tp2
                        environment.WaveLambda2 = data["Gamma2"][j]  # Gamma2
                        
                        
                        environment.WaveSeed = int(data["WaveSeed"][j])  # WaveSeed
                        environment.WaveNumberOfComponents = int(data3.iloc[7, 1]) #NwaveComp
                        vessel = model[Vessel_name]
                        vessel.InitialHeading = data["InitialHeading"][j]  # WaveInitialHeading
                        vessel.Draught = "{}".format(df)
                        vessel.InitialZ = -list_draft[list_draft.columns[1]][i]
                        vessel.ResponseStormDuration = data["ResponseStormDuration"][
                            j
                        ]  # Stormduration
                        vessel.ResponseOutputPointx[0] = data["ResponseOutputPointx[1]"][j]  # Coord X
                        vessel.ResponseOutputPointy[0] = data["ResponseOutputPointy[1]"][j]  # Coord Y
                        vessel.ResponseOutputPointz[0] = data["ResponseOutputPointz[1]"][j]  # Coord Z
                        # Creating file names
                        casenumber += 1
                        filename = "VRA{:04d}_Hs={:05.2f}_Tp={:05.2f}_{}".format(
                            casenumber, data["Hs"][j], data["Tp"][j], df
                        )
                        vessel.SaveSpectralResponseSpreadsheet(r"./01.VRA/{}.xlsx".format(filename))
                        # Savefiles if True
                        if self.Save == True:
                            model.SaveData(r"./01.VRA/{}.dat".format(filename))
                        # OpenSprectralfile

                        xl = pd.ExcelFile(r"./01.VRA/{}.xlsx".format(filename))
                        sheet = xl.sheet_names[1]
                        aux = xl.parse(sheet_name=sheet, header=6)
                        list_filename.append(filename)
                        # Extracting Results
                        if self.Acc == True:
                            list_Heave_acc.append(aux.iloc[16, 4])
                        else:
                            list_Heave_acc.append(aux.iloc[9, 4])
                        list_roll.append(aux.iloc[2, 5])
                        list_pitch.append(aux.iloc[2, 6])
                        list_rot.append((aux.iloc[2, 6] ** 2 + aux.iloc[2, 5] ** 2) ** 0.5)
                        list_hs.append(data["Hs"][j])  # Hs
                        list_tp.append(data["Tp"][j])  # Tp
                        list_dir.append(data["WaveDir"][j])  # Dir
                        list_gamma.append(data["Gamma"][j])  # Gamma
                        list_wavetype.append(data["WaveType"][j])  # Wavetype
                        list_spectralpar.append(data["WaveSpectrumParameter"][j])  # Jonswap Par
                        list_waveseed.append(int(data["WaveSeed"][j]))  # Seed
                        list_heading.append(data["InitialHeading"][j])  # Initialheading
                        list_x.append(data["ResponseOutputPointx[1]"][j])  # X
                        list_y.append(data["ResponseOutputPointy[1]"][j])  # Y
                        list_z.append(data["ResponseOutputPointz[1]"][j])  # Z
                        list_offset.append(data["Offstet[m]"][j])  # Offset
                        list_current.append(data["CurrentName"][j])  # Current
                        list_PerReturn.append(data["Return Period [years]"][j])  # ReturnPeriod
                        list_rao.append(df)
                        list_initialz.append(-list_draft[list_draft.columns[1]][i])
                        list_hs2.append(data["Hs2"][j]) #Hs2
                        list_tp2.append(data["Tp2"][j]) #Tp2
                        list_gamma2.append(data["Gamma2"][j]) #Gamma2
                        xl.close()

            else:

                data = self.data
                model = self.model
                Vessel_name = self.Vessel_name

                for j in tqdm(data.index, desc="Cases-VRA", total=len(data)):
                        # Set Environment conditions
                        environment = model.environment
                        environment.SelectedWave = "Wave1"
                        environment.WaveDirection = data["WaveDir"][j]  # WaveDir
                        environment.WaveType = data["WaveType"][j]  # WaveType
                        environment.WaveOchiHubbleParameters = data["WaveSpectrumParameter"][j]  # JonswapPar
                        environment.WaveHs1 = data["Hs"][j]  # Hs1
                        environment.Wavefm1 = 1/data["Tp"][j]  # Tp1
                        environment.WaveLambda1 = data["Gamma"][j]  # Gamma1
                        environment.WaveHs2 = data["Hs2"][j]  # Hs2
                        environment.Wavefm2 = 1/data["Tp2"][j]  # Tp2
                        environment.WaveLambda2 = data["Gamma2"][j]  # Gamma2
                        environment.WaveSeed = int(data["WaveSeed"][j])  # WaveSeed
                        environment.WaveNumberOfComponents = int(data3.iloc[7, 1])
                        vessel = model[Vessel_name]
                        vessel.InitialHeading = data["InitialHeading"][j]  # WaveInitialHeading
                        vessel.Draught = data["RAO"][j]
                        vessel.InitialZ = -data["Draft"][j]
                        vessel.ResponseStormDuration = data["ResponseStormDuration"][j]  # Stormduration
                        vessel.ResponseOutputPointx[0] = data["ResponseOutputPointx[1]"][j]  # Coord X
                        vessel.ResponseOutputPointy[0] = data["ResponseOutputPointy[1]"][j]  # Coord Y
                        vessel.ResponseOutputPointz[0] = data["ResponseOutputPointz[1]"][j]  # Coord Z
                        # Creating file names
                        casenumber += 1
                        filename = "VRA{:04d}_Hs={:05.2f}_Tp={:05.2f}_{}".format(
                            casenumber, data["Hs"][j], data["Tp"][j], vessel.Draught
                        )
                        vessel.SaveSpectralResponseSpreadsheet(r"./01.VRA/{}.xlsx".format(filename))
                        # Savefiles if True
                        if self.Save == True:
                            model.SaveData(r"./01.VRA/{}.dat".format(filename))
                        # OpenSprectralfile
                        xl = pd.ExcelFile(r"./01.VRA/{}.xlsx".format(filename))
                        sheet = xl.sheet_names[1]
                        aux = xl.parse(sheet_name=sheet, header=6)
                        list_filename.append(filename)
                        # Extracting Results
                        if self.Acc == True:
                            list_Heave_acc.append(aux.iloc[16, 4])
                        else:
                            list_Heave_acc.append(aux.iloc[9, 4])
                        list_roll.append(aux.iloc[2, 5])
                        list_pitch.append(aux.iloc[2, 6])
                        list_rot.append((aux.iloc[2, 6] ** 2 + aux.iloc[2, 5] ** 2) ** 0.5)
                        list_hs.append(data["Hs"][j])  # Hs
                        list_tp.append(data["Tp"][j])  # Tp
                        list_dir.append(data["WaveDir"][j])  # Dir
                        list_gamma.append(data["Gamma"][j])  # Gamma
                        list_wavetype.append(data["WaveType"][j])  # Wavetype
                        list_spectralpar.append(data["WaveSpectrumParameter"][j])  # Jonswap Par
                        list_waveseed.append(int(data["WaveSeed"][j]))  # Seed
                        list_heading.append(data["InitialHeading"][j])  # Initialheading
                        list_x.append(data["ResponseOutputPointx[1]"][j])  # X
                        list_y.append(data["ResponseOutputPointy[1]"][j])  # Y
                        list_z.append(data["ResponseOutputPointz[1]"][j])  # Z
                        list_offset.append(data["Offstet[m]"][j])  # Offset
                        list_current.append(data["CurrentName"][j])  # Current
                        list_PerReturn.append(data["Return Period [years]"][j])  # ReturnPeriod
                        list_rao.append(data["RAO"][j])
                        list_initialz.append(-data["Draft"][j])
                        list_hs2.append(data["Hs2"][j]) #Hs2
                        list_tp2.append(data["Tp2"][j]) #Tp2
                        list_gamma2.append(data["Gamma2"][j]) #Gamma2
                        xl.close()

            # printing
            results = pd.DataFrame(index=list_filename)
            if self.Acc == True:
                results["Z-Acceleration Amplitude [m/s2]"] = list_Heave_acc
            else:
                results["Z-Velocity Amplitude [m/s]"] = list_Heave_acc
            results["Roll Amplitude [deg]"] = list_roll
            results["Pitch Amplitude [deg]"] = list_pitch
            results["Rotation [deg]"] = list_rot
            results["Hs[m]"] = list_hs
            results["Tp[s]"] = list_tp
            results["Wave Direction[deg]"] = list_dir
            results["Gamma"] = list_gamma
            results["Wavetype"] = list_wavetype
            results["WaveSpectrumParameter"] = list_spectralpar
            results["WaveSeed"] = list_waveseed
            results["InitialHeading"] = list_heading
            results["x[m]"] = list_x
            results["y[m]"] = list_y
            results["z[m]"] = list_z
            results["Offset[m]"] = list_offset
            results["CurrenName"] = list_current
            results["Return Period"] = list_PerReturn
            results["RAO"] = list_rao
            results["Draft"] = list_initialz
            results["Hs2[m]"] = list_hs2
            results["Tp2[s]"] = list_tp2
            results["Gamma2"] = list_gamma2

        

        # Selecting Max
        aux2 = []
        d = results["Return Period"].unique()
        for i, k in enumerate(d):
            if self.Acc == True:
                a = results[results["Return Period"] == k].nlargest(
                    1, ["Z-Acceleration Amplitude [m/s2]"]
                )
                a["Criteria"] = "Acc"
            else:
                a = results[results["Return Period"] == k].nlargest(
                    1, ["Z-Velocity Amplitude [m/s]"]
                )
                a["Criteria"] = "Vel"

            b = results[results["Return Period"] == k].nlargest(1, ["Rotation [deg]"])
            b["Criteria"] = "Rot"
            # c=results[results['Return Period']==k].nlargest(1,['Pitch Amplitude [deg]'])
            # c['Criteria']='Pitch'
            # e=pd.concat([a,b,c])
            e = pd.concat([a, b])
            aux2.append(e)
        Max_Values = pd.concat(aux2)
        # Saving
        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                print("Printing VRA Results")
                # workBook.remove(workBook['VRA_Results'])
                # workBook.remove(workBook['VRA_Max'])
            except:
                print("Printing VRA Results")
            finally:
                print("Printing VRA Results")
                results.to_excel(writer, sheet_name="VRA_Results")
                Max_Values.to_excel(writer, sheet_name="VRA_Max")
        end = timer()
        print(
            f"Elapsed time - VRA : {end - start}",
        )

        
    def RunVRAGE(self, VRAmax_sheet):
        # PART2_____________PART2
        """
        Requirement: Run  LoadCasesVRA
        Runs time screening
        Args:
            VRAmax_sheet: Name of the spreadsheet with Max_Values of VRA"""

        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        start2 = timer()
        file = self.ExcelFile
        data3 = self.data3

        stage1 = data3.iloc[1, 1]
        stage2 = data3.iloc[1, 2]
        WaveComp = data3.iloc[7, 1]

        if os.path.isdir("02.GE") == False:
            os.mkdir("02.GE")
        list_Acc = []
        list_T = []
        list_Roll = []
        list_Pitch = []
        list_criteria = []
        list_case = []
        model = self.model
        # vessel=self.vessel

        if self.spect==0:

            for j, i in tqdm(enumerate(Max_Values.index), desc="GE Running", total=len(Max_Values)):
                general = model.general
                general.DynamicsEnabled='Yes' 
                general.StageDuration[0] = stage1
                general.StageDuration[1] = stage2
                environment = model.environment
                environment.SelectedWave = "Wave1"
                environment.WaveOriginX = 0
                environment.WaveOriginY = 0
                environment.WaveDirection = Max_Values["Wave Direction[deg]"][j]  # WaveDir
                environment.WaveHs = Max_Values["Hs[m]"][j]  # Hs
                environment.WaveType = Max_Values["Wavetype"][j]  # WaveType
                environment.WaveJONSWAPParameters = Max_Values["WaveSpectrumParameter"][
                    j
                ]  # Jonswapparm
                environment.WaveGamma = Max_Values["Gamma"][j]  # Gamma
                environment.WaveTp = Max_Values["Tp[s]"][j]  # Tp
                environment.WaveNumberOfComponents=int(WaveComp) #NComp
                environment.WaveSeed = Max_Values["WaveSeed"][j]  # Seed
                vessel = model[self.Vessel_name]
                vessel.InitialHeading = Max_Values["InitialHeading"][j]  # vesselHeading
                x = Max_Values["x[m]"][j]  # Coord x
                y = Max_Values["y[m]"][j]  # Coord y
                z = Max_Values["z[m]"][j]  # Coord z

                vessel.Draught = Max_Values["RAO"][j]  # Select RAO
                vessel.InitialZ = Max_Values["Draft"][j]  # Select draft

                model.RunSimulation()
                per = 1  # Stage of simulation
                obj = OrcFxAPI.oeVessel(x, y, z)
                if self.Acc == True:
                    varName = ["Gz acceleration", "Roll", "Pitch"]
                    stats = vessel.LinkedStatistics(
                        varName, period=per, objectExtra=obj
                    )  # Vessel time history
                    if Max_Values['Criteria'][j]=='Acc':
                        query = stats.Query("Gz acceleration", "Roll")
                        list_Acc.append(query.ValueAtMax)  # Max Acc
                        list_Acc.append(query.ValueAtMin)  # Min Acc
                        list_T.append(query.TimeOfMax)  # time of Max acc
                        list_criteria.append("MaxAcc")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min acc
                        list_criteria.append("MinAcc")
                        list_case.append(i)
                        list_Roll.append(query.LinkedValueAtMax)  # Roll linked max Acc
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Acc
                        query = stats.Query("Gz acceleration", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Acc
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Acc
                    ##
                    elif Max_Values['Criteria'][j]=='Rot':
                        query = stats.Query("Roll", "Gz acceleration")
                        list_Roll.append(query.ValueAtMax)  # Max Roll
                        list_Roll.append(query.ValueAtMin)  # Min Roll
                        list_T.append(query.TimeOfMax)  # time of Max Roll
                        list_criteria.append("MaxRoll")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Roll
                        list_criteria.append("MinRoll")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Roll
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Roll
                        query = stats.Query("Roll", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Roll
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Roll
                    ##
                        query = stats.Query("Pitch", "Gz acceleration")
                        list_Pitch.append(query.ValueAtMax)  # Max Pitch
                        list_Pitch.append(query.ValueAtMin)  # Min Pitch
                        list_T.append(query.TimeOfMax)  # time of Max Picth
                        list_criteria.append("MaxPitch")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Pitch
                        list_criteria.append("MinPitch")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Pitch
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Pitch
                        query = stats.Query("Pitch", "Roll")
                        list_Roll.append(query.LinkedValueAtMax)  # Pitch linked max Pitch
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Pitch
                    # File_name
                    model.SaveSimulation(r"./02.GE/{}.sim".format(i))
                    results2 = pd.DataFrame(index=list_case)
                    results2["Z-Acceleration Amplitude [m/s2]"] = list_Acc
                    results2["Roll Amplitude [deg]"] = list_Roll
                    results2["Pitch Amplitude [deg]"] = list_Pitch
                    results2["Time"] = list_T
                    results2["Criteria"] = list_criteria
                else:
                    varName = ["Gz velocity", "Roll", "Pitch"]
                    stats = vessel.LinkedStatistics(
                        varName, period=per, objectExtra=obj
                    )  # Vessel time history
                    if Max_Values['Criteria'][j]=='Vel':
                        query = stats.Query("Gz velocity", "Roll")
                        list_Acc.append(query.ValueAtMax)  # Max Acc
                        list_Acc.append(query.ValueAtMin)  # Min Acc
                        list_T.append(query.TimeOfMax)  # time of Max acc
                        list_criteria.append("MaxVel")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min acc
                        list_criteria.append("MinVel")
                        list_case.append(i)
                        list_Roll.append(query.LinkedValueAtMax)  # Roll linked max Acc
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Acc
                        query = stats.Query("Gz velocity", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Acc
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Acc
                    ##
                    elif Max_Values['Criteria'][j]=='Rot':
                        query = stats.Query("Roll", "Gz velocity")
                        list_Roll.append(query.ValueAtMax)  # Max Roll
                        list_Roll.append(query.ValueAtMin)  # Min Roll
                        list_T.append(query.TimeOfMax)  # time of Max Roll
                        list_criteria.append("MaxRoll")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Roll
                        list_criteria.append("MinRoll")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Roll
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Roll
                        query = stats.Query("Roll", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Roll
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Roll
                        ##
                        query = stats.Query("Pitch", "Gz velocity")
                        list_Pitch.append(query.ValueAtMax)  # Max Pitch
                        list_Pitch.append(query.ValueAtMin)  # Min Pitch
                        list_T.append(query.TimeOfMax)  # time of Max Picth
                        list_criteria.append("MaxPitch")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Pitch
                        list_criteria.append("MinPitch")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Pitch
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Pitch
                        query = stats.Query("Pitch", "Roll")
                        list_Roll.append(query.LinkedValueAtMax)  # Pitch linked max Pitch
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Pitch
                    # File_name
                    model.SaveSimulation(r"./02.GE/{}.sim".format(i))
                    results2 = pd.DataFrame(index=list_case)
                    results2["Z-Velocity Amplitude [m/s]"] = list_Acc
                    results2["Roll Amplitude [deg]"] = list_Roll
                    results2["Pitch Amplitude [deg]"] = list_Pitch
                    results2["Time"] = list_T
                    results2["Criteria"] = list_criteria

        if self.spect==1:

            for j, i in tqdm(enumerate(Max_Values.index), desc="GE Running", total=len(Max_Values)):
                general = model.general
                general.DynamicsEnabled='Yes' 
                general.StageDuration[0] = stage1
                general.StageDuration[1] = stage2
                environment = model.environment
                environment.SelectedWave = "Wave1"
                environment.WaveOriginX = 0
                environment.WaveOriginY = 0
                environment.WaveDirection = Max_Values["Wave Direction[deg]"][j]  # WaveDir
                environment.WaveType = Max_Values["Wavetype"][j]  # WaveType
                environment.WaveOchiHubbleParameters = Max_Values["WaveSpectrumParameter"][j]  # Jonswapparm
                environment.WaveHs1 = Max_Values["Hs[m]"][j]  # Hs
                environment.Wavefm1 = 1/Max_Values["Tp[s]"][j]  # Tp
                environment.WaveLambda1 = Max_Values["Gamma"][j]  # Gamma
                environment.WaveHs2 = Max_Values["Hs2[m]"][j]  # Hs2
                environment.Wavefm2 = 1/Max_Values["Tp2[s]"][j]  # Tp2
                environment.WaveLambda2 = Max_Values["Gamma2"][j]  # Gamma2
                environment.WaveNumberOfComponents=int(WaveComp) #NComp
                environment.WaveSeed = Max_Values["WaveSeed"][j]  # Seed
                vessel = model[self.Vessel_name]
                vessel.InitialHeading = Max_Values["InitialHeading"][j]  # vesselHeading
                x = Max_Values["x[m]"][j]  # Coord x
                y = Max_Values["y[m]"][j]  # Coord y
                z = Max_Values["z[m]"][j]  # Coord z

                vessel.Draught = Max_Values["RAO"][j]  # Select RAO
                vessel.InitialZ = Max_Values["Draft"][j]  # Select draft

                model.RunSimulation()
                per = 1  # Stage of simulation
                obj = OrcFxAPI.oeVessel(x, y, z)
                if self.Acc == True:
                    varName = ["Gz acceleration", "Roll", "Pitch"]
                    stats = vessel.LinkedStatistics(
                        varName, period=per, objectExtra=obj
                    )  # Vessel time history
                    if Max_Values['Criteria'][j]=='Acc':
                        query = stats.Query("Gz acceleration", "Roll")
                        list_Acc.append(query.ValueAtMax)  # Max Acc
                        list_Acc.append(query.ValueAtMin)  # Min Acc
                        list_T.append(query.TimeOfMax)  # time of Max acc
                        list_criteria.append("MaxAcc")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min acc
                        list_criteria.append("MinAcc")
                        list_case.append(i)
                        list_Roll.append(query.LinkedValueAtMax)  # Roll linked max Acc
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Acc
                        query = stats.Query("Gz acceleration", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Acc
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Acc
                    ##
                    elif Max_Values['Criteria'][j]=='Rot':
                        query = stats.Query("Roll", "Gz acceleration")
                        list_Roll.append(query.ValueAtMax)  # Max Roll
                        list_Roll.append(query.ValueAtMin)  # Min Roll
                        list_T.append(query.TimeOfMax)  # time of Max Roll
                        list_criteria.append("MaxRoll")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Roll
                        list_criteria.append("MinRoll")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Roll
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Roll
                        query = stats.Query("Roll", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Roll
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Roll
                    ##
                        query = stats.Query("Pitch", "Gz acceleration")
                        list_Pitch.append(query.ValueAtMax)  # Max Pitch
                        list_Pitch.append(query.ValueAtMin)  # Min Pitch
                        list_T.append(query.TimeOfMax)  # time of Max Picth
                        list_criteria.append("MaxPitch")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Pitch
                        list_criteria.append("MinPitch")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Pitch
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Pitch
                        query = stats.Query("Pitch", "Roll")
                        list_Roll.append(query.LinkedValueAtMax)  # Pitch linked max Pitch
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Pitch
                    # File_name
                    model.SaveSimulation(r"./02.GE/{}.sim".format(i))
                    results2 = pd.DataFrame(index=list_case)
                    results2["Z-Acceleration Amplitude [m/s2]"] = list_Acc
                    results2["Roll Amplitude [deg]"] = list_Roll
                    results2["Pitch Amplitude [deg]"] = list_Pitch
                    results2["Time"] = list_T
                    results2["Criteria"] = list_criteria
                else:
                    varName = ["Gz velocity", "Roll", "Pitch"]
                    stats = vessel.LinkedStatistics(
                        varName, period=per, objectExtra=obj
                    )  # Vessel time history
                    if Max_Values['Criteria'][j]=='Vel':
                        query = stats.Query("Gz velocity", "Roll")
                        list_Acc.append(query.ValueAtMax)  # Max Acc
                        list_Acc.append(query.ValueAtMin)  # Min Acc
                        list_T.append(query.TimeOfMax)  # time of Max acc
                        list_criteria.append("MaxVel")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min acc
                        list_criteria.append("MinVel")
                        list_case.append(i)
                        list_Roll.append(query.LinkedValueAtMax)  # Roll linked max Acc
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Acc
                        query = stats.Query("Gz velocity", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Acc
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Acc
                    ##
                    elif Max_Values['Criteria'][j]=='Rot':
                        query = stats.Query("Roll", "Gz velocity")
                        list_Roll.append(query.ValueAtMax)  # Max Roll
                        list_Roll.append(query.ValueAtMin)  # Min Roll
                        list_T.append(query.TimeOfMax)  # time of Max Roll
                        list_criteria.append("MaxRoll")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Roll
                        list_criteria.append("MinRoll")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Roll
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Roll
                        query = stats.Query("Roll", "Pitch")
                        list_Pitch.append(query.LinkedValueAtMax)  # Pitch linked max Roll
                        list_Pitch.append(query.LinkedValueAtMin)  # Pitch linked min Roll
                        ##
                        query = stats.Query("Pitch", "Gz velocity")
                        list_Pitch.append(query.ValueAtMax)  # Max Pitch
                        list_Pitch.append(query.ValueAtMin)  # Min Pitch
                        list_T.append(query.TimeOfMax)  # time of Max Picth
                        list_criteria.append("MaxPitch")
                        list_case.append(i)
                        list_T.append(query.TimeOfMin)  # time of Min Pitch
                        list_criteria.append("MinPitch")
                        list_case.append(i)
                        list_Acc.append(query.LinkedValueAtMax)  # Acc linked max Pitch
                        list_Acc.append(query.LinkedValueAtMin)  # Acc linked min Pitch
                        query = stats.Query("Pitch", "Roll")
                        list_Roll.append(query.LinkedValueAtMax)  # Pitch linked max Pitch
                        list_Roll.append(query.LinkedValueAtMin)  # Roll linked min Pitch
                    # File_name
                    model.SaveSimulation(r"./02.GE/{}.sim".format(i))
                    results2 = pd.DataFrame(index=list_case)
                    results2["Z-Velocity Amplitude [m/s]"] = list_Acc
                    results2["Roll Amplitude [deg]"] = list_Roll
                    results2["Pitch Amplitude [deg]"] = list_Pitch
                    results2["Time"] = list_T
                    results2["Criteria"] = list_criteria
        # Saving
        with pd.ExcelWriter(
            file, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GE_Results'])
                print("Printing GE Results")
            except:
                print("Printing GE Results")
            finally:
                print("Printing GE Results")
                results2.to_excel(writer, sheet_name="GE_Results")
        end2 = timer()
        print(
            f"Elapsed time - GE : {end2 - start2}",
        )

    def CreateCasesSheetGA(self, GE_Resultsheet, VRAmax_sheet):

        """
        Creates spreadsheet for Load Cases for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet:
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults"""
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)

        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_hs2=[]
        list_tp2=[]
        list_gamma2=[]
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
        ##

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir
        if self.spect==0:

            for j, i in enumerate(results2.index):
                auxF = Max_Values.loc[i]
                if auxF.ndim > 1:
                    auxF = auxF.iloc[0]

                for m, n in enumerate(setdir.index):
                    LC += 1
                    list_casename.append("GA{:04d}_{}_{}_{}".format(LC, i, results2["Criteria"][j], n))
                    list_time.append(results2["Time"][j] - 0.5 * Extstage2)
                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                    list_OffsetData.append(auxF.loc["Offset[m]"])
                    list_Xoffset.append(
                        auxF.loc["Offset[m]"] * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                    )
                    list_Yoffset.append(
                        auxF.loc["Offset[m]"] * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                    )
                    list_hs.append(auxF.loc["Hs[m]"])
                    list_gamma.append(auxF.loc["Gamma"])
                    list_tp.append(auxF.loc["Tp[s]"])
                    list_dir.append(auxF.loc["Wave Direction[deg]"])
                    list_rao.append(auxF.loc["RAO"])
                    list_initialz.append(auxF.loc["Draft"])
                    list_heading.append(auxF.loc["InitialHeading"])
                    list_wavetype.append(auxF.loc["Wavetype"])
                    list_spectralpar.append(auxF.loc["WaveSpectrumParameter"])
                    list_x.append(auxF.loc["x[m]"])
                    list_y.append(auxF.loc["y[m]"])
                    list_z.append(auxF.loc["z[m]"])
                    list_waveseed.append(auxF.loc["WaveSeed"])
                    list_current.append(auxF.loc["CurrenName"])

            vet = zip(
                list_hs,
                list_tp,
                list_dir,
                list_gamma,
                list_wavetype,
                list_spectralpar,
                list_waveseed,
                list_heading,
                list_x,
                list_y,
                list_z,
                list_rao,
                list_initialz,
                list_current,
                list_time,
            )
            result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

            result3["CurrentDir[deg]"] = list_OffSetDirData
            result3["Offset[m]"] = list_OffsetData
            result3["OffsetDir[deg]"] = list_OffSetDirData
            result3["XOffset[m]"] = list_Xoffset
            result3["YOffset[m]"] = list_Yoffset
            result3["XWaveOrigin[m]"] = list_Xoffset
            result3["YWaveOrigin[m]"] = list_Yoffset
            result3["NwaveComp"]=self.data3.iloc[7,1]

        if self.spect==1:
            for j, i in enumerate(results2.index):
                auxF = Max_Values.loc[i]
                if auxF.ndim > 1:
                    auxF = auxF.iloc[0]

                for m, n in enumerate(setdir.index):
                    LC += 1
                    list_casename.append("GA{:04d}_{}_{}_{}".format(LC, i, results2["Criteria"][j], n))
                    list_time.append(results2["Time"][j] - 0.5 * Extstage2)
                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                    list_OffsetData.append(auxF.loc["Offset[m]"])
                    list_Xoffset.append(
                        auxF.loc["Offset[m]"] * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                    )
                    list_Yoffset.append(
                        auxF.loc["Offset[m]"] * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                    )
                    list_hs.append(auxF.loc["Hs[m]"])
                    list_gamma.append(auxF.loc["Gamma"])
                    list_tp.append(auxF.loc["Tp[s]"])
                    list_dir.append(auxF.loc["Wave Direction[deg]"])
                    list_rao.append(auxF.loc["RAO"])
                    list_initialz.append(auxF.loc["Draft"])
                    list_heading.append(auxF.loc["InitialHeading"])
                    list_wavetype.append(auxF.loc["Wavetype"])
                    list_spectralpar.append(auxF.loc["WaveSpectrumParameter"])
                    list_x.append(auxF.loc["x[m]"])
                    list_y.append(auxF.loc["y[m]"])
                    list_z.append(auxF.loc["z[m]"])
                    list_waveseed.append(auxF.loc["WaveSeed"])
                    list_current.append(auxF.loc["CurrenName"])
                    list_hs2.append(auxF.loc["Hs2[m]"])
                    list_tp2.append(auxF.loc["Tp2[s]"])
                    list_gamma2.append(auxF.loc["Gamma2"])

            vet = zip(
                list_hs,
                list_tp,
                list_dir,
                list_gamma,
                list_wavetype,
                list_spectralpar,
                list_waveseed,
                list_heading,
                list_x,
                list_y,
                list_z,
                list_rao,
                list_initialz,
                list_current,
                list_time,
            )
            result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

            result3["CurrentDir[deg]"] = list_OffSetDirData
            result3["Offset[m]"] = list_OffsetData
            result3["OffsetDir[deg]"] = list_OffSetDirData
            result3["XOffset[m]"] = list_Xoffset
            result3["YOffset[m]"] = list_Yoffset
            result3["XWaveOrigin[m]"] = list_Xoffset
            result3["YWaveOrigin[m]"] = list_Yoffset
            result3["NwaveComp"]=self.data3.iloc[7,1]
            result3["Hs2[m]"] = list_hs2
            result3["Tp2[s]"] = list_tp2
            result3["Gamma2"] = list_gamma2


        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_Cases")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases part-3 : {end3 - start3}",
        )

    def CreateWaveCasesYellowTail(self, GE_Resultsheet, VRAmax_sheet):
        """
        Creates spreadsheet for Load Cases for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet:
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults"""
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)

        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_fluid = []
        list_casetype = []
        list_heel = []
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
        ##
        # Offset-Interpolation
        off_1yr_int = [
            [
                47.2,
                53.6,
                69.9,
                71.7,
                84.65,
                97.6,
                101.2,
                104.7,
                118.9,
                133.1,
                134.3,
                136,
                166.2,
                174.4,
                191.7,
                196.6,
                199,
                240.2,
                246.5,
                257.1,
            ],
            [
                12,
                15.4,
                29.6,
                31.5,
                44.4,
                57.3,
                57.8,
                58.1,
                55.3,
                51.4,
                51.3,
                50.8,
                47.3,
                47.3,
                23.9,
                35.8,
                40.2,
                43.5,
                45.2,
                45.8,
            ],
        ]

        off_1yr_dam = [
            [
                32.6,
                41.7,
                63.3,
                65.6,
                80.05,
                94.5,
                98.3,
                102.2,
                117.2,
                132.1,
                133.3,
                135,
                167.3,
                176.2,
                184.7,
                191.6,
                194.7,
                236.4,
                243,
                254.1,
            ],
            [
                15.6,
                19.1,
                34.7,
                36.9,
                51.35,
                65.8,
                66.1,
                66.4,
                62.5,
                57.4,
                57.2,
                56.6,
                52.5,
                52.7,
                30,
                43.8,
                48.9,
                50,
                51.7,
                51.8,
            ],
        ]

        off_100yrs_int = [
            [
                6.1,
                61.4,
                64.6,
                70.95,
                102.6,
                102.7,
                104.3,
                104.9,
                105.7,
                122.7,
                137.2,
                138,
                138.1,
                152.2,
                153.1,
                154.2,
                168.6,
                193.1,
                195.4,
                196.1,
                258.7,
                263.1,
                353.8,
            ],
            [
                14.2,
                27.8,
                30,
                66.45,
                103.2,
                102.9,
                106.3,
                106.9,
                107.4,
                104.7,
                100.3,
                100.2,
                100.1,
                98.4,
                98.4,
                98.2,
                99,
                46,
                50.2,
                51.1,
                98.4,
                99.2,
                13.9,
            ],
        ]

        off_100yrs_dam = [
            [
                33.5,
                37.9,
                61.05,
                95.8,
                96.2,
                98,
                98.6,
                99.6,
                118.4,
                134.7,
                135.6,
                136,
                152.3,
                153.4,
                154.6,
                171.6,
                180,
                180,
                252.6,
                257.5,
                272.9,
                342.2,
                348.4,
            ],
            [
                46.2,
                47.7,
                92.65,
                137.6,
                138.1,
                142.1,
                142.7,
                143.2,
                135,
                126.2,
                126,
                126,
                122.6,
                122.8,
                122.5,
                124.8,
                82.4,
                81.3,
                128.5,
                128.5,
                76.2,
                41.2,
                40.3,
            ],
        ]

        def f1(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_1yr_int[0], fp=off_1yr_int[1], period=360)
            return y

        def f2(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_1yr_dam[0], fp=off_1yr_dam[1], period=360)
            return y

        def f3(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_100yrs_int[0], fp=off_100yrs_int[1], period=360)
            return y

        def f4(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_100yrs_dam[0], fp=off_100yrs_dam[1], period=360)
            return y

        def offset(x, l, cond):
            """
            x: Offset direction ccw from E
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if l == "1 YRP" or l == "1 YR":
                if cond == "int":
                    y = f1(x)
                else:
                    y = f2(x)

            else:
                if cond == "int":
                    y = f3(x)
                else:
                    y = f4(x)
            return y
        
        def offset2(cond):
            """
            CaLaculates the onmidirectional offset
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if cond=='dam':
               y=143.2
                    
            else :
                y=107.4
                
            return y

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

        def Rec_Cur(a):
            if a >= 105 and a < 130:
                label = "BIN1_95%"
            elif a >= 130 and a < 185:
                label = "BIN2_95%"
            elif (a >= 185 and a <= 360) or (a >= 0 and a < 105):
                label = "BIN3_95%"
            else:
                label = "ERRO"
            return label

        list_cases = ["Recurrent", "Extreme", "Abnormal", "Survival", "SIT"]
        # Recurrent Cases
        for cases in list_cases:
            if cases == "Recurrent":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_Rec_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=offset2('int')
                            list_OffsetData.append(off)

                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                            list_fluid.append("MeanOperating")
                            list_casetype.append(cases)
                            list_heel.append(0)

            if cases == "Extreme":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["100 YRP", "100 YR"])]

                list_Ofluid = ["Empty", "MaxOperating", "DeadOil"]

                for g in list_Ofluid:

                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                list_fluid.append(g)
                                list_casetype.append(cases)
                                list_heel.append(0)

            if cases == "Abnormal":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_Abn_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                            #)
                            off=offset2('dam')
                            list_OffsetData.append(off)
                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                            list_fluid.append("Associateddensity")
                            list_casetype.append(cases)
                            list_heel.append(0)

            if cases == "Survival":

                ##Def Heel by Draft
                def assheel(name):
                    a = name.split("_")[0]
                    if a == "Min":
                        h = [-8.3, 13.3]
                    elif a == "LC50":
                        h = [-8.3, 13.3]
                    elif a == "Max":
                        h = [-8.3, 13.3]
                    return h

                Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP", "1 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        heel = assheel(Max_r["RAO"][j])
                        for h in heel:
                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Sur_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=58.10
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                list_fluid.append("MeanOperating")
                                list_casetype.append(cases)
                                list_heel.append(h)

                Max_r = Max_Values[Max_Values["Return Period"].isin(["1000 YRP", "1000 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_Sur_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=offset2('dam')
                            list_OffsetData.append(off)
                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                            list_fluid.append("MeanOperating")
                            list_casetype.append(cases)
                            list_heel.append(0)

            if cases == "SIT":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_SIT_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=58.10
                            list_OffsetData.append(off)
                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                            list_fluid.append("SeaWater")
                            list_casetype.append(cases)
                            list_heel.append(0)

        vet = zip(
            list_hs,
            list_tp,
            list_dir,
            list_gamma,
            list_wavetype,
            list_spectralpar,
            list_waveseed,
            list_heading,
            list_x,
            list_y,
            list_z,
            list_rao,
            list_initialz,
            list_current,
            list_time,
        )
        result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

        result3["CurrentDir[deg]"] = list_OffSetDirData
        result3["Offset[m]"] = list_OffsetData
        result3["OffsetDir[deg]"] = list_OffSetDirData
        result3["XOffset[m]"] = list_Xoffset
        result3["YOffset[m]"] = list_Yoffset
        result3["XWaveOrigin[m]"] = list_Xoffset
        result3["YWaveOrigin[m]"] = list_Yoffset
        result3["Fluid"] = list_fluid
        result3["CaseType"] = list_casetype
        result3["Heel"] = list_heel
        result3["NwaveComp"]=self.data3.iloc[7,1]

        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_Cases_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )
        
    def CreateCurrentCasesYellowTail(self, GE_Resultsheet, VRAmax_sheet, current_sheet):
        """
        Creates spreadsheet for Load Cases for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet: Results of VRA
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults
            current_sheet; Name of the spreadsheet with Max_Values of Current screening
            """
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)
        
        sheetname5 = current_sheet
        Cur_res = pd.read_excel(self.ExcelFile, sheet_name=sheetname5, header=0, index_col=0)
        
        Cur_res=Cur_res.drop_duplicates(['Current_profile','Return_period'], keep='first')


        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_fluid = []
        list_casetype = []
        list_heel = []
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
        ##
        # Offset-Interpolation
        off_1yr_int = [
            [
                47.2,
                53.6,
                69.9,
                71.7,
                84.65,
                97.6,
                101.2,
                104.7,
                118.9,
                133.1,
                134.3,
                136,
                166.2,
                174.4,
                191.7,
                196.6,
                199,
                240.2,
                246.5,
                257.1,
            ],
            [
                12,
                15.4,
                29.6,
                31.5,
                44.4,
                57.3,
                57.8,
                58.1,
                55.3,
                51.4,
                51.3,
                50.8,
                47.3,
                47.3,
                23.9,
                35.8,
                40.2,
                43.5,
                45.2,
                45.8,
            ],
        ]

        off_1yr_dam = [
            [
                32.6,
                41.7,
                63.3,
                65.6,
                80.05,
                94.5,
                98.3,
                102.2,
                117.2,
                132.1,
                133.3,
                135,
                167.3,
                176.2,
                184.7,
                191.6,
                194.7,
                236.4,
                243,
                254.1,
            ],
            [
                15.6,
                19.1,
                34.7,
                36.9,
                51.35,
                65.8,
                66.1,
                66.4,
                62.5,
                57.4,
                57.2,
                56.6,
                52.5,
                52.7,
                30,
                43.8,
                48.9,
                50,
                51.7,
                51.8,
            ],
        ]

        off_100yrs_int = [
            [
                6.1,
                61.4,
                64.6,
                70.95,
                102.6,
                102.7,
                104.3,
                104.9,
                105.7,
                122.7,
                137.2,
                138,
                138.1,
                152.2,
                153.1,
                154.2,
                168.6,
                193.1,
                195.4,
                196.1,
                258.7,
                263.1,
                353.8,
            ],
            [
                14.2,
                27.8,
                30,
                66.45,
                103.2,
                102.9,
                106.3,
                106.9,
                107.4,
                104.7,
                100.3,
                100.2,
                100.1,
                98.4,
                98.4,
                98.2,
                99,
                46,
                50.2,
                51.1,
                98.4,
                99.2,
                13.9,
            ],
        ]

        off_100yrs_dam = [
            [
                33.5,
                37.9,
                61.05,
                95.8,
                96.2,
                98,
                98.6,
                99.6,
                118.4,
                134.7,
                135.6,
                136,
                152.3,
                153.4,
                154.6,
                171.6,
                180,
                180,
                252.6,
                257.5,
                272.9,
                342.2,
                348.4,
            ],
            [
                46.2,
                47.7,
                92.65,
                137.6,
                138.1,
                142.1,
                142.7,
                143.2,
                135,
                126.2,
                126,
                126,
                122.6,
                122.8,
                122.5,
                124.8,
                82.4,
                81.3,
                128.5,
                128.5,
                76.2,
                41.2,
                40.3,
            ],
        ]

        def f1(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_1yr_int[0], fp=off_1yr_int[1], period=360)
            return y

        def f2(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_1yr_dam[0], fp=off_1yr_dam[1], period=360)
            return y

        def f3(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_100yrs_int[0], fp=off_100yrs_int[1], period=360)
            return y

        def f4(r):
            # r should be a vertor [x1,x2,...]
            y = np.interp(x=r, xp=off_100yrs_dam[0], fp=off_100yrs_dam[1], period=360)
            return y

        def offset(x, l, cond):
            """
            x: Offset direction ccw from E
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if l == "1 YRP" or l == "1 YR":
                if cond == "int":
                    y = f1(x)
                else:
                    y = f2(x)

            else:
                if cond == "int":
                    y = f3(x)
                else:
                    y = f4(x)
            return y
        
        def offset2(cond):
            """
            CaLaculates the onmidirectional offset
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if cond=='dam':
               y=143.2
                    
            else :
                y=107.4
                
            return y

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

        def Rec_Cur(a):
            if a >= 105 and a < 130:
                label = "BIN1_95%"
            elif a >= 130 and a < 185:
                label = "BIN2_95%"
            elif (a >= 185 and a <= 360) or (a >= 0 and a < 105):
                label = "BIN3_95%"
            else:
                label = "ERRO"
            return label
        
        

        list_cases = ["Recurrent", "Extreme", "Abnormal", "Survival", "SIT"]
        # Recurrent Cases
        for cases in list_cases:
            if cases == "Recurrent":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                Max_C = Cur_res[Cur_res['Return_period'].isin(['10 YRP', '10 YR'])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        

                        for m, n in enumerate(setdir.index):
                            for r in range(len(Max_C.index)):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_C_Rec_{}_{}_{}_{}_{}".format(LC, i,Max_C["Criteria"][r] ,Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)
    
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Max_C["Current_profile"][r])
                                list_fluid.append("MeanOperating")
                                list_casetype.append(cases)
                                list_heel.append(0)

            if cases == "Extreme":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                Max_C = Cur_res[Cur_res['Return_period'].isin(['100 YRP', '100 YR'])]

                list_Ofluid = ["Empty", "MaxOperating", "DeadOil"]

                for g in list_Ofluid:

                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                for r in range(len(Max_C.index)):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_C_Ext_{}_{}_{}_{}_{}".format(LC, i,Max_C["Criteria"][r] ,Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('int')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(Max_C["Current_profile"][r])
                                    list_fluid.append(g)
                                    list_casetype.append(cases)
                                    list_heel.append(0)

            if cases == "Abnormal":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                Max_C = Cur_res[Cur_res['Return_period'].isin(['10 YRP', '10 YR'])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            for r in range(len(Max_C.index)):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_C_Abn_{}_{}_{}_{}_{}".format(LC, i,Max_C["Criteria"][r] ,Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                                #)
                                off=offset2('dam')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Max_C["Current_profile"][r])
                                list_fluid.append("Associateddensity")
                                list_casetype.append(cases)
                                list_heel.append(0)

            if cases == "Survival":

                ##Def Heel by Draft
                def assheel(name):
                    a = name.split("_")[0]
                    if a == "Min":
                        h = [-8.3, 13.3]
                    elif a == "LC50":
                        h = [-8.3, 13.3]
                    elif a == "Max":
                        h = [-8.3, 13.3]
                    return h

                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                Max_C = Cur_res[Cur_res['Return_period'].isin(['1 YRP', '1 YR'])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        heel = assheel(Max_r["RAO"][j])
                        for h in heel:
                            for m, n in enumerate(setdir.index):
                                for r in range(len(Max_C.index)):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_C_Sur_{}_{}_{}_{}_{}".format(LC, i,Max_C["Criteria"][r] ,Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=58.1
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(Max_C["Current_profile"][r])
                                    list_fluid.append("MeanOperating")
                                    list_casetype.append(cases)
                                    list_heel.append(h)

                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                Max_C = Cur_res[Cur_res['Return_period'].isin(['1000 YRP', '1000 YR'])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            for r in range(len(Max_C.index)):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_C_Sur_{}_{}_{}_{}_{}".format(LC, i,Max_C["Criteria"][r] ,Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('dam')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Max_C["Current_profile"][r])
                                list_fluid.append("MeanOperating")
                                list_casetype.append(cases)
                                list_heel.append(0)

            if cases == "SIT":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_C_SIT_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=58.10
                            list_OffsetData.append(off)
                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append("BIN2_95%")
                            list_fluid.append("SeaWater")
                            list_casetype.append(cases)
                            list_heel.append(0)

        vet = zip(
            list_hs,
            list_tp,
            list_dir,
            list_gamma,
            list_wavetype,
            list_spectralpar,
            list_waveseed,
            list_heading,
            list_x,
            list_y,
            list_z,
            list_rao,
            list_initialz,
            list_current,
            list_time,
        )
        result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

        result3["CurrentDir[deg]"] = list_OffSetDirData
        result3["Offset[m]"] = list_OffsetData
        result3["OffsetDir[deg]"] = list_OffSetDirData
        result3["XOffset[m]"] = list_Xoffset
        result3["YOffset[m]"] = list_Yoffset
        result3["XWaveOrigin[m]"] = list_Xoffset
        result3["YWaveOrigin[m]"] = list_Yoffset
        result3["Fluid"] = list_fluid
        result3["CaseType"] = list_casetype
        result3["Heel"] = list_heel
        result3["NwaveComp"]=self.data3.iloc[7,1]

        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_Cases_CUR_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )

    def CreateWaveCasesKizomba(self, GE_Resultsheet, VRAmax_sheet):
        """
        Creates spreadsheet for Load Cases for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet:
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults"""
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)

        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_fluid = []
        list_casetype = []
        list_heel = []
        list_hs2=[]
        list_tp2=[]
        list_gamma2=[]
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
       
        
        def offset2(cond):
            """
            CaLaculates the onmidirectional offset
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if cond=='dam':
               y=60.96
                    
            else :
                y=50.8
                
            return y

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

        def Rec_Cur(a):
            if a >= 105 and a < 130:
                label = "AssC"
            elif a >= 130 and a < 185:
                label = "AssC"
            elif (a >= 185 and a <= 360) or (a >= 0 and a < 105):
                label = "AssC"
            else:
                label = "ERRO"
            return label

        list_cases = ["Normal", "Extreme", 'Abnormal-1YR', 'Abnormal–10YR']
        # Recurrent Cases
        
        if self.spect==0:

            for cases in list_cases:
                if cases == "Normal":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Nor_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)

                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)
                                

                if cases == "Extreme":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["100 YRP", "100 YR"])]

                    list_Ofluid = ["Operating","Empty"]

                    for g in list_Ofluid:

                        for j, i in enumerate(Max_r.index):
                            auxF = results2.loc[i]

                            for k in range(len(auxF.index)):

                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('int')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                    list_fluid.append(g)
                                    list_casetype.append(cases)
                                    list_heel.append(0)

                if cases == "Abnormal-1YR":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP", "1 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Abn1yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                                #)
                                off=offset2('dam')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)

                if cases == "Abnormal–10YR":

                    ##Def Heel by Draft
                    def assheel(name):
                        h= [-6, 6]
                        return h

                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            heel = assheel(Max_r["RAO"][j])
                            for h in heel:
                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Abn10yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('dam')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                    list_fluid.append("Operating")
                                    list_casetype.append(cases)
                                    list_heel.append(h)

            vet = zip(
                list_hs,
                list_tp,
                list_dir,
                list_gamma,
                list_wavetype,
                list_spectralpar,
                list_waveseed,
                list_heading,
                list_x,
                list_y,
                list_z,
                list_rao,
                list_initialz,
                list_current,
                list_time,
            )
            result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

            result3["CurrentDir[deg]"] = list_OffSetDirData
            result3["Offset[m]"] = list_OffsetData
            result3["OffsetDir[deg]"] = list_OffSetDirData
            result3["XOffset[m]"] = list_Xoffset
            result3["YOffset[m]"] = list_Yoffset
            result3["XWaveOrigin[m]"] = list_Xoffset
            result3["YWaveOrigin[m]"] = list_Yoffset
            result3["Fluid"] = list_fluid
            result3["CaseType"] = list_casetype
            result3["Heel"] = list_heel
            result3["NwaveComp"]=self.data3.iloc[7,1]

        if self.spect==1:

            for cases in list_cases:
                if cases == "Normal":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Nor_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)

                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)
                                list_hs2.append(Max_r["Hs2[m]"][j])
                                list_tp2.append(Max_r["Tp2[s]"][j])
                                list_gamma2.append(Max_r["Gamma2"][j])

                if cases == "Extreme":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["100 YRP", "100 YR"])]

                    list_Ofluid = ["Operating","Empty"]

                    for g in list_Ofluid:

                        for j, i in enumerate(Max_r.index):
                            auxF = results2.loc[i]

                            for k in range(len(auxF.index)):

                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('int')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                    list_fluid.append(g)
                                    list_casetype.append(cases)
                                    list_heel.append(0)
                                    list_hs2.append(Max_r["Hs2[m]"][j])
                                    list_tp2.append(Max_r["Tp2[s]"][j])
                                    list_gamma2.append(Max_r["Gamma2"][j])

                if cases == "Abnormal-1YR":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP", "1 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Abn1yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                                #)
                                off=offset2('dam')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)
                                list_hs2.append(Max_r["Hs2[m]"][j])
                                list_tp2.append(Max_r["Tp2[s]"][j])
                                list_gamma2.append(Max_r["Gamma2"][j])

                if cases == "Abnormal–10YR":

                    ##Def Heel by Draft
                    def assheel(name):
                        h= [-6, 6]
                        return h

                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            heel = assheel(Max_r["RAO"][j])
                            for h in heel:
                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Abn10yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('dam')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(Rec_Cur(setdir[setdir.columns[0]][m]))
                                    list_fluid.append("Operating")
                                    list_casetype.append(cases)
                                    list_heel.append(h)
                                    list_hs2.append(Max_r["Hs2[m]"][j])
                                    list_tp2.append(Max_r["Tp2[s]"][j])
                                    list_gamma2.append(Max_r["Gamma2"][j])

            vet = zip(
                list_hs,
                list_tp,
                list_dir,
                list_gamma,
                list_wavetype,
                list_spectralpar,
                list_waveseed,
                list_heading,
                list_x,
                list_y,
                list_z,
                list_rao,
                list_initialz,
                list_current,
                list_time,
            )
            result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

            result3["CurrentDir[deg]"] = list_OffSetDirData
            result3["Offset[m]"] = list_OffsetData
            result3["OffsetDir[deg]"] = list_OffSetDirData
            result3["XOffset[m]"] = list_Xoffset
            result3["YOffset[m]"] = list_Yoffset
            result3["XWaveOrigin[m]"] = list_Xoffset
            result3["YWaveOrigin[m]"] = list_Yoffset
            result3["Fluid"] = list_fluid
            result3["CaseType"] = list_casetype
            result3["Heel"] = list_heel
            result3["NwaveComp"]=self.data3.iloc[7,1]
            result3["Hs2[m]"] = list_hs2
            result3["Tp2[s]"] = list_tp2
            result3["Gamma2"]= list_gamma2


        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_Cases__Wave_Kizomba")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )

    def CreateCurrentCasesKizomba(self, GE_Resultsheet, VRAmax_sheet):
        """
        Creates spreadsheet for Load Cases for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet:
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults"""
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)

        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_fluid = []
        list_casetype = []
        list_heel = []
        list_hs2=[]
        list_tp2=[]
        list_gamma2=[]
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
       
        
        def offset2(cond):
            """
            CaLaculates the onmidirectional offset
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if cond=='dam':
               y=60.96
                    
            else :
                y=50.8
                
            return y

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

        def Rec_Cur(a):
            if a >= 105 and a < 130:
                label = "AssC"
            elif a >= 130 and a < 185:
                label = "AssC"
            elif (a >= 185 and a <= 360) or (a >= 0 and a < 105):
                label = "AssC"
            else:
                label = "ERRO"
            return label

        list_cases = ["Normal", "Extreme", 'Abnormal-1YR', 'Abnormal–10YR']
        # Recurrent Cases
        
        if self.spect==0:

            for cases in list_cases:
                if cases == "Normal":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Nor_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)

                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append('10yr_Cur')
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)
                                

                if cases == "Extreme":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["100 YRP", "100 YR"])]

                    list_Ofluid = ["Operating","Empty"]

                    for g in list_Ofluid:

                        for j, i in enumerate(Max_r.index):
                            auxF = results2.loc[i]

                            for k in range(len(auxF.index)):

                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('int')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append('100yr_Cur')
                                    list_fluid.append(g)
                                    list_casetype.append(cases)
                                    list_heel.append(0)

                if cases == "Abnormal-1YR":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP", "1 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Abn1yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                                #)
                                off=offset2('dam')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append('1yr_Cur')
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)

                if cases == "Abnormal–10YR":

                    ##Def Heel by Draft
                    def assheel(name):
                        h= [-6, 6]
                        return h

                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            heel = assheel(Max_r["RAO"][j])
                            for h in heel:
                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Abn10yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('dam')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append('10yr_Cur')
                                    list_fluid.append("Operating")
                                    list_casetype.append(cases)
                                    list_heel.append(h)

            vet = zip(
                list_hs,
                list_tp,
                list_dir,
                list_gamma,
                list_wavetype,
                list_spectralpar,
                list_waveseed,
                list_heading,
                list_x,
                list_y,
                list_z,
                list_rao,
                list_initialz,
                list_current,
                list_time,
            )
            result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

            result3["CurrentDir[deg]"] = list_OffSetDirData
            result3["Offset[m]"] = list_OffsetData
            result3["OffsetDir[deg]"] = list_OffSetDirData
            result3["XOffset[m]"] = list_Xoffset
            result3["YOffset[m]"] = list_Yoffset
            result3["XWaveOrigin[m]"] = list_Xoffset
            result3["YWaveOrigin[m]"] = list_Yoffset
            result3["Fluid"] = list_fluid
            result3["CaseType"] = list_casetype
            result3["Heel"] = list_heel
            result3["NwaveComp"]=self.data3.iloc[7,1]

        if self.spect==1:

            for cases in list_cases:
                if cases == "Normal":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Nor_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)

                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append('10yr_Cur')
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)
                                list_hs2.append(Max_r["Hs2[m]"][j])
                                list_tp2.append(Max_r["Tp2[s]"][j])
                                list_gamma2.append(Max_r["Gamma2"][j])

                if cases == "Extreme":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["100 YRP", "100 YR"])]

                    list_Ofluid = ["Operating","Empty"]

                    for g in list_Ofluid:

                        for j, i in enumerate(Max_r.index):
                            auxF = results2.loc[i]

                            for k in range(len(auxF.index)):

                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('int')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append('100yr_Cur')
                                    list_fluid.append(g)
                                    list_casetype.append(cases)
                                    list_heel.append(0)
                                    list_hs2.append(Max_r["Hs2[m]"][j])
                                    list_tp2.append(Max_r["Tp2[s]"][j])
                                    list_gamma2.append(Max_r["Gamma2"][j])

                if cases == "Abnormal-1YR":
                    Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP", "1 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Abn1yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                                #)
                                off=offset2('dam')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append('1yr_Cur')
                                list_fluid.append("Operating")
                                list_casetype.append(cases)
                                list_heel.append(0)
                                list_hs2.append(Max_r["Hs2[m]"][j])
                                list_tp2.append(Max_r["Tp2[s]"][j])
                                list_gamma2.append(Max_r["Gamma2"][j])

                if cases == "Abnormal–10YR":

                    ##Def Heel by Draft
                    def assheel(name):
                        h= [-6, 6]
                        return h

                    Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):
                            heel = assheel(Max_r["RAO"][j])
                            for h in heel:
                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Abn10yr_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    off=offset2('dam')
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append('10yr_Cur')
                                    list_fluid.append("Operating")
                                    list_casetype.append(cases)
                                    list_heel.append(h)
                                    list_hs2.append(Max_r["Hs2[m]"][j])
                                    list_tp2.append(Max_r["Tp2[s]"][j])
                                    list_gamma2.append(Max_r["Gamma2"][j])

            vet = zip(
                list_hs,
                list_tp,
                list_dir,
                list_gamma,
                list_wavetype,
                list_spectralpar,
                list_waveseed,
                list_heading,
                list_x,
                list_y,
                list_z,
                list_rao,
                list_initialz,
                list_current,
                list_time,
            )
            result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

            result3["CurrentDir[deg]"] = list_OffSetDirData
            result3["Offset[m]"] = list_OffsetData
            result3["OffsetDir[deg]"] = list_OffSetDirData
            result3["XOffset[m]"] = list_Xoffset
            result3["YOffset[m]"] = list_Yoffset
            result3["XWaveOrigin[m]"] = list_Xoffset
            result3["YWaveOrigin[m]"] = list_Yoffset
            result3["Fluid"] = list_fluid
            result3["CaseType"] = list_casetype
            result3["Heel"] = list_heel
            result3["NwaveComp"]=self.data3.iloc[7,1]
            result3["Hs2[m]"] = list_hs2
            result3["Tp2[s]"] = list_tp2
            result3["Gamma2"]= list_gamma2


        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_Cases_Cur_Kizomba")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )

    def CreateWaveCasesAkerYellowTail(self, GE_Resultsheet, VRAmax_sheet):
        """
        Creates spreadsheet for Load Wave Cases Aker YellowTail for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet:
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults"""
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)

        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_fluid = []
        list_casetype = []
        list_heel = []
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
    

        def offset2(cond):
            """
            CaLaculates the onmidirectional offset
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if cond=='dam':
               y=117
                    
            else :
                y=156
                
            return y

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

    
        list_cases = ["Normal", "Extreme", "Abnormal", "Survival", "SIT"]
        # Normal Cases
        for cases in list_cases:
            if cases == "Normal":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["10 YRP", "10 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_Nor_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=offset2('int')
                            list_OffsetData.append(off)

                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append("AssC")
                            list_fluid.append("MeanOperating")
                            list_casetype.append(cases)
                            list_heel.append(0)

            if cases == "Extreme":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["100 YRP", "100 YR"])]

                list_Ofluid = ["MeanOperating"]

                for g in list_Ofluid:

                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append("AssC")
                                list_fluid.append(g)
                                list_casetype.append(cases)
                                list_heel.append(0)

           

            if cases == "Abnormal":

                ##Def Heel by Draft
                def assheel(a):
                    if a == "1 YRP" or "1 YR":
                        h = [-10, 10]
                    elif a == "10 YRP" or "10 YR":
                        h = [-6, 6]
                    elif a == "100 YRP" or "100 YR":
                        h = [0]
                    return h

                Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP", "1 YR","10 YRP","10 YR","100 YRP","100 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        heel = assheel(Max_r["Return Period"][j])
                        for h in heel:
                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Abn_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                aa='dam' if Max_r["Return Period"][j]==("100 YRP" or "100 YR") else 'int'
                                off=offset2(aa)
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append("AssC")
                                list_fluid.append("MeanOperating")
                                list_casetype.append(cases)
                                list_heel.append(h)

            if cases == "Survival":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["1000 YRP","1000 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_Sur_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "dam"
                            #)
                            off=offset2('int')
                            list_OffsetData.append(off)
                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append('AssC')
                            list_fluid.append("MeanOperating")
                            list_casetype.append(cases)
                            list_heel.append(0)   

            if cases == "SIT":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["1 YRP","1 YR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_SIT_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=offset2('int')
                            list_OffsetData.append(off)
                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append('AssC')
                            list_fluid.append("SeaWater")
                            list_casetype.append(cases)
                            list_heel.append(0)

        vet = zip(
            list_hs,
            list_tp,
            list_dir,
            list_gamma,
            list_wavetype,
            list_spectralpar,
            list_waveseed,
            list_heading,
            list_x,
            list_y,
            list_z,
            list_rao,
            list_initialz,
            list_current,
            list_time,
        )
        result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

        result3["CurrentDir[deg]"] = list_OffSetDirData
        result3["Offset[m]"] = list_OffsetData
        result3["OffsetDir[deg]"] = list_OffSetDirData
        result3["XOffset[m]"] = list_Xoffset
        result3["YOffset[m]"] = list_Yoffset
        result3["XWaveOrigin[m]"] = list_Xoffset
        result3["YWaveOrigin[m]"] = list_Yoffset
        result3["Fluid"] = list_fluid
        result3["CaseType"] = list_casetype
        result3["Heel"] = list_heel
        result3["NwaveComp"]=self.data3.iloc[7,1]

        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_WaveCases_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )

    def CreateCurrentCasesAkerYellowTail(self, GE_Resultsheet, VRAmax_sheet):
        """
        Creates spreadsheet for Load Current Cases Aker YellowTail for GA (Far, Near,Cros and Transversal )
        Args:
            VRAmax_sheet:
            GE_Resultsheet: Name of the spreadsheet with Max_Values of GEresults"""
        start3 = timer()
        sheetname3 = VRAmax_sheet
        Max_Values = pd.read_excel(self.ExcelFile, sheet_name=sheetname3, header=0, index_col=0)

        sheetname4 = GE_Resultsheet
        results2 = pd.read_excel(self.ExcelFile, sheet_name=sheetname4, header=0, index_col=0)

        riser_az = self.data3.iloc[1, 0]
        Extstage1 = self.data3.iloc[1, 3]
        Extstage2 = self.data3.iloc[1, 4]

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_time = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        ##
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_current = []
        list_fluid = []
        list_casetype = []
        list_heel = []
        ##
        list_par = [
            "Hs[m]",
            "Tp[s]",
            "Wave Direction[deg]",
            "Gamma",
            "Wavetype",
            "WaveSpectrumParameter",
            "WaveSeed",
            "InitialHeading",
            "x[m]",
            "y[m]",
            "z[m]",
            "RAO",
            "Draft",
            "CurrenName",
            "SimulationTimeOrigin",
        ]
    

        def offset2(cond):
            """
            CaLaculates the onmidirectional offset
            l: Label of the Wave (1 YRP, 10YRP, etc)
            cond: Mooring condition-'int' or 'dam'
            """
            if cond=='dam':
               y=117
                    
            else :
                y=156
                
            return y

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

    
        list_cases = ["Normal", "Extreme", "Abnormal"]
        # Normal Cases
        for cases in list_cases:
            if cases == "Normal":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        

                        for m, n in enumerate(setdir.index):
                            LC += 1
                            list_casename.append(
                                "{:04d}_GA_Nor_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                            )
                            list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                            list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                            #off = offset(
                                #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                            #)
                            off=offset2('int')
                            list_OffsetData.append(off)

                            list_Xoffset.append(
                                off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_Yoffset.append(
                                off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                            )
                            list_hs.append(Max_r["Hs[m]"][j])
                            list_gamma.append(Max_r["Gamma"][j])
                            list_tp.append(Max_r["Tp[s]"][j])
                            list_dir.append(Max_r["Wave Direction[deg]"][j])
                            list_rao.append(Max_r["RAO"][j])
                            list_initialz.append(Max_r["Draft"][j])
                            list_heading.append(Max_r["InitialHeading"][j])
                            list_wavetype.append(Max_r["Wavetype"][j])
                            list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                            list_x.append(Max_r["x[m]"][j])
                            list_y.append(Max_r["y[m]"][j])
                            list_z.append(Max_r["z[m]"][j])
                            list_waveseed.append(Max_r["WaveSeed"][j])
                            list_current.append("10yrC")
                            list_fluid.append("MeanOperating")
                            list_casetype.append(cases)
                            list_heel.append(0)

            if cases == "Extreme":
                Max_r = Max_Values[Max_Values["Return Period"].isin(["CUR"])]

                list_Ofluid = ["MeanOperating"]

                for g in list_Ofluid:

                    for j, i in enumerate(Max_r.index):
                        auxF = results2.loc[i]

                        for k in range(len(auxF.index)):

                            for m, n in enumerate(setdir.index):
                                LC += 1
                                list_casename.append(
                                    "{:04d}_GA_Ext_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                )
                                list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                #off = offset(
                                    #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                #)
                                off=offset2('int')
                                list_OffsetData.append(off)
                                list_Xoffset.append(
                                    off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_Yoffset.append(
                                    off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                )
                                list_hs.append(Max_r["Hs[m]"][j])
                                list_gamma.append(Max_r["Gamma"][j])
                                list_tp.append(Max_r["Tp[s]"][j])
                                list_dir.append(Max_r["Wave Direction[deg]"][j])
                                list_rao.append(Max_r["RAO"][j])
                                list_initialz.append(Max_r["Draft"][j])
                                list_heading.append(Max_r["InitialHeading"][j])
                                list_wavetype.append(Max_r["Wavetype"][j])
                                list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                list_x.append(Max_r["x[m]"][j])
                                list_y.append(Max_r["y[m]"][j])
                                list_z.append(Max_r["z[m]"][j])
                                list_waveseed.append(Max_r["WaveSeed"][j])
                                list_current.append("100yrC")
                                list_fluid.append(g)
                                list_casetype.append(cases)
                                list_heel.append(0)

           

            if cases == "Abnormal":

                ##Def Heel by Draft
                def assheel(name):
                    if a == '1yrC':
                        h = [-10, 10]
                    elif a == '10yrC':
                        h = [-6, 6]
                    elif a == '100yrC':
                        h = [0]
                    return h
                Currents=['1yrC', '10yrC', '100yrC']

                Max_r = Max_Values[Max_Values["Return Period"].isin(['CUR'])]
                for j, i in enumerate(Max_r.index):
                    auxF = results2.loc[i]

                    for k in range(len(auxF.index)):
                        for a in Currents:
                            heel = assheel(a)
                            for h in heel:
                                for m, n in enumerate(setdir.index):
                                    LC += 1
                                    list_casename.append(
                                        "{:04d}_GA_Abn_{}_{}_{}_{}".format(LC, i, Max_r["Criteria"][j],auxF["Criteria"][k] ,n)
                                    )
                                    list_time.append(auxF["Time"][k] - 0.5 * Extstage2)
                                    list_OffSetDirData.append(setdir[setdir.columns[0]][m])
                                    #off = offset(
                                        #setdir[setdir.columns[0]][m], Max_r["Return Period"][j], "int"
                                    #)
                                    aa='dam' if a=='100yrC'else 'int'
                                    off=offset2(aa)
                                    list_OffsetData.append(off)
                                    list_Xoffset.append(
                                        off * np.cos(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_Yoffset.append(
                                        off * np.sin(np.deg2rad(setdir[setdir.columns[0]][m]))
                                    )
                                    list_hs.append(Max_r["Hs[m]"][j])
                                    list_gamma.append(Max_r["Gamma"][j])
                                    list_tp.append(Max_r["Tp[s]"][j])
                                    list_dir.append(Max_r["Wave Direction[deg]"][j])
                                    list_rao.append(Max_r["RAO"][j])
                                    list_initialz.append(Max_r["Draft"][j])
                                    list_heading.append(Max_r["InitialHeading"][j])
                                    list_wavetype.append(Max_r["Wavetype"][j])
                                    list_spectralpar.append(Max_r["WaveSpectrumParameter"][j])
                                    list_x.append(Max_r["x[m]"][j])
                                    list_y.append(Max_r["y[m]"][j])
                                    list_z.append(Max_r["z[m]"][j])
                                    list_waveseed.append(Max_r["WaveSeed"][j])
                                    list_current.append(a)
                                    list_fluid.append("MeanOperating")
                                    list_casetype.append(cases)
                                    list_heel.append(h)

           
        vet = zip(
            list_hs,
            list_tp,
            list_dir,
            list_gamma,
            list_wavetype,
            list_spectralpar,
            list_waveseed,
            list_heading,
            list_x,
            list_y,
            list_z,
            list_rao,
            list_initialz,
            list_current,
            list_time,
        )
        result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

        result3["CurrentDir[deg]"] = list_OffSetDirData
        result3["Offset[m]"] = list_OffsetData
        result3["OffsetDir[deg]"] = list_OffSetDirData
        result3["XOffset[m]"] = list_Xoffset
        result3["YOffset[m]"] = list_Yoffset
        result3["XWaveOrigin[m]"] = list_Xoffset
        result3["YWaveOrigin[m]"] = list_Yoffset
        result3["Fluid"] = list_fluid
        result3["CaseType"] = list_casetype
        result3["Heel"] = list_heel
        result3["NwaveComp"]=self.data3.iloc[7,1]

        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing GA_Cases")
            except:
                print("Printing GA_Cases")
            finally:
                print("Printing GA_Cases")
                result3.to_excel(writer, sheet_name="GA_CurCases_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )


    def CasesVRATime(self):
        """
        Creates VRA Orcaflex load case and save in a folder for Time screening methodology
        Requirement: Run  LoadCasesVRA
        """
        if os.path.isdir("01.VRA_Time2") == False:
            os.mkdir("01.VRA_Time2")
        start = timer()

        list_PerReturn = []
        list_hs = []
        list_dir = []
        list_tp = []
        list_rao = []
        list_gamma = []
        list_wavetype = []
        list_spectralpar = []
        list_initialz = []
        list_heading = []
        list_x = []
        list_y = []
        list_z = []
        list_waveseed = []
        list_offset = []
        list_current = []
        list_case = []
        list_criteria = []

        casenumber = 0

        list_draft = self.list_draft
        data = self.data
        model = self.model
        Vessel_name = self.Vessel_name

        stage1 = self.data3.iloc[1, 1]
        stage2 = self.data3.iloc[1, 2]
        Nvar = self.data3.iloc[3, 4]
        WaveComp=self.data.iloc[7,1]
        varName = []
        for k, n in enumerate(range(Nvar)):
            varName.append(self.data3.iloc[4 + k, 4])

        for i, df in tqdm(
            enumerate(list_draft.iloc[:, 0]), desc="Draft-VRA", total=len(list_draft)
        ):
            for j in tqdm(data.index, desc="Cases-VRA_Time", total=len(data)):
                general = model.general
                general.StageDuration[0] = stage1
                general.StageDuration[1] = stage2
                # Set Environment conditions
                environment = model.environment
                environment.SelectedWave = "Wave1"
                environment.WaveDirection = data["WaveDir"][j]  # WaveDir
                environment.WaveHs = data["Hs"][j]  # Hs
                environment.WaveOriginX = 0
                environment.WaveOriginY = 0
                environment.WaveType = data["WaveType"][j]  # WaveType
                environment.WaveJONSWAPParameters = data["WaveSpectrumParameter"][j]  # JonswapPar
                environment.WaveGamma = data["Gamma"][j]  # Gamma
                environment.WaveTp = data["Tp"][j]  # Tp
                environment.WaveSeed = int(data["WaveSeed"][j])  # WaveSeed
                environment.WaveNumberOfComponents=WaveComp #NComp
                vessel = model[Vessel_name]
                vessel.InitialHeading = data["InitialHeading"][j]  # WaveInitialHeading
                vessel.Draught = "{}".format(df)
                vessel.InitialZ = -list_draft[list_draft.columns[1]][i]
                vessel.ResponseOutputPointx[0] = data["ResponseOutputPointx[1]"][j]  # Save Coord X
                vessel.ResponseOutputPointy[0] = data["ResponseOutputPointy[1]"][j]  # Save Coord Y
                vessel.ResponseOutputPointz[0] = data["ResponseOutputPointz[1]"][j]  # Save Coord Z
                # Creating file names
                casenumber += 1
                filename = "VRA_Time_{:04d}_Hs={:05.2f}_Tp={:05.2f}_{}".format(
                    casenumber, environment.WaveHs, environment.WaveTp, df
                )
                # vessel.SaveSpectralResponseSpreadsheet(r'./01.VRA/{}.xlsx'.format(filename))
                # Savefiles if True
                model.SaveData(r"./01.VRA_Time2/{}.dat".format(filename))

                for k, n in enumerate(range(Nvar)):
                    list_criteria.extend(
                        ["Max_{}".format(varName), "Min_{}".format(varName)]
                    )  # Criteria
                    list_case.extend([filename, filename])
                    list_hs.extend([data["Hs"][j], data["Hs"][j]])  # Hs
                    list_tp.extend([data["Tp"][j], data["Tp"][j]])  # Tp
                    list_dir.extend([data["WaveDir"][j], data["WaveDir"][j]])  # Dir
                    list_gamma.extend([data["Gamma"][j], data["Gamma"][j]])  # Gamma
                    list_wavetype.extend([data["WaveType"][j], data["WaveType"][j]])  # Wavetype
                    list_spectralpar.extend(
                        [data["WaveSpectrumParameter"][j], data["WaveSpectrumParameter"][j]]
                    )  # Jonswap Par
                    list_waveseed.extend(
                        [int(data["WaveSeed"][j]), int(data["WaveSeed"][j])]
                    )  # Seed
                    list_heading.extend(
                        [data["InitialHeading"][j], data["InitialHeading"][j]]
                    )  # Initialheading
                    list_x.extend(
                        [data["ResponseOutputPointx[1]"][j], data["ResponseOutputPointx[1]"][j]]
                    )  # X
                    list_y.extend(
                        [data["ResponseOutputPointy[1]"][j], data["ResponseOutputPointy[1]"][j]]
                    )  # Y
                    list_z.extend(
                        [data["ResponseOutputPointz[1]"][j], data["ResponseOutputPointz[1]"][j]]
                    )  # Z
                    list_offset.extend([data["Offstet[m]"][j], data["Offstet[m]"][j]])  # Offset
                    list_current.extend(
                        [data["CurrentName"][j], data["CurrentName"][j]]
                    )  # Current
                    list_PerReturn.extend(
                        [data["Return Period [years]"][j], data["Return Period [years]"][j]]
                    )  # ReturnPeriod
                    list_rao.extend([df, df])  # RAO
                    list_initialz.extend(
                        [
                            -list_draft[list_draft.columns[1]][i],
                            -list_draft[list_draft.columns[1]][i],
                        ]
                    )  # -DRAFT
        self.parcial3 = list_case
        # printing
        results = pd.DataFrame(index=list_case)
        results["Hs[m]"] = list_hs
        results["Tp[s]"] = list_tp
        results["Wave Direction[deg]"] = list_dir
        results["Gamma"] = list_gamma
        results["Wavetype"] = list_wavetype
        results["WaveSpectrumParameter"] = list_spectralpar
        results["WaveSeed"] = list_waveseed
        results["NwaveComp"] = WaveComp
        results["Vessel heading"] = list_heading
        results["x[m]"] = list_x
        results["y[m]"] = list_y
        results["z[m]"] = list_z
        results["Offset[m]"] = list_offset
        results["CurrenName"] = list_current
        results["Return Period"] = list_PerReturn
        results["RAO"] = list_rao
        results["Draft"] = list_initialz

        self.VRAResultsheet = "VRA_Time_Results"

        # Saving
        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                print("Printing VRA Results")
                # workBook.remove(workBook['VRA_Results'])
                # workBook.remove(workBook['VRA_Max'])
            except:
                print("Printing VRA Time Results")
            finally:
                print("Printing VRA Time Results")
                results.to_excel(writer, sheet_name="{}".format(self.VRAResultsheet))
        end = timer()
        print(
            f"Elapsed time - VRA_Time : {end - start}",
        )

    def ProcessVRATime(self, folder):
        """
        Processes VRA Orcaflex .sim files for Time screeening methodology
        Requirement: Run  LoadCasesVRA and .sim files
        """
        start = timer()

        Vessel_name = self.Vessel_name
        Nvar = self.data3.iloc[3, 4]
        list_Res = []
        list_T = []

        varName = []

        for k, n in enumerate(range(Nvar)):
            varName.append(self.data3.iloc[4 + k, 4])

        list_Res = [[] for _ in range(Nvar)]

        files = glob.glob("{}/*.sim".format(folder))
        model = OrcFxAPI.Model()

        for file in tqdm(files, desc="Processing VRA", total=len(files)):
            model.LoadSimulation(file)
            vessel = model[Vessel_name]
            x = vessel.ResponseOutputPointx[0]  # Save Coord X
            y = vessel.ResponseOutputPointy[0]  # Save Coord Y
            z = vessel.ResponseOutputPointz[0]  # Save Coord Z
            per = 1  # Stage of simulation
            obj = OrcFxAPI.oeVessel(x, y, z)
            stats = vessel.LinkedStatistics(
                varName, period=per, objectExtra=obj
            )  # Vessel time history
            for k, n in enumerate(range(Nvar)):

                # varName=self.data3.iloc[4+k,4]

                for m, s in enumerate(range(Nvar)):
                    query = stats.Query(varName[k], varName[m])
                    if k == m:
                        list_Res[m].extend([query.ValueAtMax, query.ValueAtMin])  # Max and Min
                    else:
                        list_Res[m].extend([query.LinkedValueAtMax, query.LinkedValueAtMin])

                list_T.extend([query.TimeOfMax, query.TimeOfMin])  # time of Max and Min

        results = pd.read_excel(
            self.ExcelFile, sheet_name=self.VRAResultsheet, header=0, index_col=0
        )
        for n, m in enumerate(varName):
            name = m
            results["{}".format(name)] = list_Res[n]
        results["Time"] = list_T
        # Selecting Max
        aux2 = []

        d = results["Return Period"].unique()
        list_Var = ["Sea surface Z", "Elevation"]
        for i, k in enumerate(d):
            name = []
            res = []
            for n, m in enumerate(varName):
                if m in [k for k in list_Var]:
                    name.append(m)
                    res.append(
                        results[results["Return Period"] == k].nlargest(1, ["{}".format(m)])
                    )
                else:
                    name.extend([m, m])
                    res.append(
                        results[results["Return Period"] == k].nlargest(1, ["{}".format(m)])
                    )
                    res.append(
                        results[results["Return Period"] == k].nsmallest(1, ["{}".format(m)])
                    )
            e = pd.concat(res)
            e["Criteria"] = name
            aux2.append(e)
        Max_Values = pd.concat(aux2)
        # Saving
        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                print("Printing VRA Results")
                # workBook.remove(workBook['VRA_Results'])
                # workBook.remove(workBook['VRA_Max'])
            except:
                print("Printing VRA Time Results")
            finally:
                print("Printing VRA Time Results")
                results.to_excel(writer, sheet_name="VRA_Time_Results")
                Max_Values.to_excel(writer, sheet_name="VRA_Time_Max")
        end = timer()
        print(
            f"Elapsed time - VRA_Time : {end - start}",
        )

    @staticmethod
    def _ThreadedRunModel(ij, list_draft, data3, data, Vessel_name, orcafile):
        def _runModel(ij, list_draft, data3, data, Vessel_name, orcafile):
            i = ij[0]
            j = ij[1]
            stage1 = data3.iloc[1, 1]
            stage2 = data3.iloc[1, 2]
            Nvar = data3.iloc[3, 4]
            WaveComp = data3.iloc[7, 1]
            varName = []
            for k, n in enumerate(range(Nvar)):
                varName.append(data3.iloc[4 + k, 4])
            result_lists = defaultdict(list)
            result_lists["Res"] = [[] for _ in range(Nvar)]
            new_model = OrcFxAPI.Model(orcafile)
            new_model.general.StageDuration[0] = stage1
            new_model.general.StageDuration[1] = stage2
            # Set Environment conditions
            new_model.environment.SelectedWave = "Wave1"
            new_model.environment.WaveDirection = data["WaveDir"][j]  # WaveDir
            new_model.environment.WaveHs = data["Hs"][j]  # Hs
            new_model.environment.WaveOriginX = 0
            new_model.environment.WaveOriginY = 0
            new_model.environment.WaveType = data["WaveType"][j]  # WaveType
            new_model.environment.WaveJONSWAPParameters = data["WaveSpectrumParameter"][
                j
            ]  # JonswapPar
            new_model.environment.WaveGamma = data["Gamma"][j]  # Gamma
            new_model.environment.WaveTp = data["Tp"][j]  # Tp
            new_model.environment.WaveSeed = int(data["WaveSeed"][j])  # WaveSeed
            new_model.environment.WaveNumberOfComponents=int(WaveComp) #Number of wave Comp
            new_model[Vessel_name].InitialHeading = data["InitialHeading"][j]  # WaveInitialHeading
            new_model[Vessel_name].Draught = "{}".format(list_draft.iloc[i, 0])
            new_model[Vessel_name].InitialZ = -list_draft[list_draft.columns[1]][i]
            # Creating file names
            casenumber = int(i * len(data) + j + 1)
            filename = "VRA_Time_{:04d}_Hs={:05.2f}_Tp={:05.2f}_{}".format(
                casenumber,
                new_model.environment.WaveHs,
                new_model.environment.WaveTp,
                list_draft.iloc[i, 0],
            )
            # vessel.SaveSpectralResponseSpreadsheet(r'./01.VRA/{}.xlsx'.format(filename))
            # Savefiles if True
            new_model.SaveData(r"./01.VRA_Time/{}.dat".format(filename))
            new_model.RunSimulation()
            per = 1  # Stage of simulation
            x = data["ResponseOutputPointx[1]"][j]  # Coord x
            y = data["ResponseOutputPointy[1]"][j]  # Coord y
            z = data["ResponseOutputPointz[1]"][j]  # Coord z
            obj = OrcFxAPI.oeVessel(x, y, z)
            stats = new_model[Vessel_name].LinkedStatistics(
                varName, period=per, objectExtra=obj
            )  # Vessel time history
            for k, n in enumerate(range(Nvar)):
                for m, s in enumerate(range(Nvar)):
                    query = stats.Query(varName[k], varName[m])
                    if k == m:
                        result_lists["Res"][m].extend(
                            [query.ValueAtMax, query.ValueAtMin]
                        )  # Max and Min
                    else:
                        result_lists["Res"][m].extend(
                            [query.LinkedValueAtMax, query.LinkedValueAtMin]
                        )

                result_lists["T"].extend([query.TimeOfMax, query.TimeOfMin])  # time of Max and Min
                result_lists["criteria"].extend(
                    ["Max_{}".format(varName), "Min_{}".format(varName)]
                )  # Criteria
                result_lists["case"].extend([filename, filename])
                result_lists["hs"].extend([data["Hs"][j], data["Hs"][j]])  # Hs
                result_lists["tp"].extend([data["Tp"][j], data["Tp"][j]])  # Tp
                result_lists["dir"].extend([data["WaveDir"][j], data["WaveDir"][j]])  # Dir
                result_lists["gamma"].extend([data["Gamma"][j], data["Gamma"][j]])  # Gamma
                result_lists["wavetype"].extend(
                    [data["WaveType"][j], data["WaveType"][j]]
                )  # Wavetype
                result_lists["spectralpar"].extend(
                    [data["WaveSpectrumParameter"][j], data["WaveSpectrumParameter"][j]]
                )  # Jonswap Par
                result_lists["waveseed"].extend(
                    [int(data["WaveSeed"][j]), int(data["WaveSeed"][j])]
                )  # Seed
                result_lists["heading"].extend(
                    [data["InitialHeading"][j], data["InitialHeading"][j]]
                )  # Initialheading
                result_lists["x"].extend(
                    [data["ResponseOutputPointx[1]"][j], data["ResponseOutputPointx[1]"][j]]
                )  # X
                result_lists["y"].extend(
                    [data["ResponseOutputPointy[1]"][j], data["ResponseOutputPointy[1]"][j]]
                )  # Y
                result_lists["z"].extend(
                    [data["ResponseOutputPointz[1]"][j], data["ResponseOutputPointz[1]"][j]]
                )  # Z
                result_lists["offset"].extend(
                    [data["Offstet[m]"][j], data["Offstet[m]"][j]]
                )  # Offset
                result_lists["current"].extend(
                    [data["CurrentName"][j], data["CurrentName"][j]]
                )  # Current
                result_lists["returnperiod"].extend(
                    [data["Return Period [years]"][j], data["Return Period [years]"][j]]
                )  # ReturnPeriod
                result_lists["rao"].extend(
                    [list_draft.loc[i, "RAO"], list_draft.loc[i, "RAO"]]
                )  # RAO
                
                result_lists["initialz"].extend(
                    [
                        -list_draft[list_draft.columns[1]][i],
                        -list_draft[list_draft.columns[1]][i],
                    ]
                )  # -DRAFT
                
                result_lists["NwaveComp"].extend(
                    [WaveComp, WaveComp]
                )
            
            return result_lists

        with ThreadingPool(2) as tpool:
            result_lists = tpool.map(_runModel, ij, list_draft, data3, data, Vessel_name, orcafile)
        return result_lists

    def VRATime(self, savefiles=True):
        """
        Requirement: Run  LoadCasesVRA
        Runs time VRA Screnning
        Args:
        savefiles:Save orcaflex files"""
        if os.path.isdir("01.VRA_Time") == False:
            os.mkdir("01.VRA_Time")
        start = timer()

        list_draft = self.list_draft
        data = self.data
        Nvar = self.data3.iloc[3, 4]
       
        varName = []

        for k, n in enumerate(range(Nvar)):
            varName.append(self.data3.iloc[4 + k, 4])

        ij = []
        for i, df in enumerate(list_draft.iloc[:, 0]):
            for j in data.index:
                ij.append((i, j))
        subs = 6
        ij = [ij[i : min([len(ij), i + subs])] for i in range(0, len(ij), subs)]
        ts = time.time()
        with ProcessingPool(subs) as pool:
            result_dicts = pool.map(
                self._ThreadedRunModel,
                ij,
                [[list_draft] * len(ij_i) for ij_i in ij],
                [[self.data3] * len(ij_i) for ij_i in ij],
                [[data] * len(ij_i) for ij_i in ij],
                [[self.Vessel_name] * len(ij_i) for ij_i in ij],
                [[self.orcafile] * len(ij_i) for ij_i in ij],
            )
        te = time.time()
        result_dicts = [item for r in result_dicts for item in r]
        print(f"Sim Time: {te-ts}")
        # printing
        results = pd.DataFrame(index=[item for r in result_dicts for item in r["case"]])
        for n, m in enumerate(varName):
            name = m
            results["{}".format(name)] = [item for r in result_dicts for item in r["Res"][0]]
        results["Time"] = [item for r in result_dicts for item in r["T"]]
        results["Hs[m]"] = [item for r in result_dicts for item in r["hs"]]
        results["Tp[s]"] = [item for r in result_dicts for item in r["tp"]]
        results["Wave Direction[deg]"] = [item for r in result_dicts for item in r["dir"]]
        results["Gamma"] = [item for r in result_dicts for item in r["gamma"]]
        results["Wavetype"] = [item for r in result_dicts for item in r["wavetype"]]
        results["WaveSpectrumParameter"] = [
            item for r in result_dicts for item in r["spectralpar"]
        ]
        results["WaveSeed"] = [item for r in result_dicts for item in r["waveseed"]]
        results["NwaveComp"] = [item for r in result_dicts for item in r["NwaveComp"]]
        results["Vessel heading"] = [item for r in result_dicts for item in r["heading"]]
        results["x[m]"] = [item for r in result_dicts for item in r["x"]]
        results["y[m]"] = [item for r in result_dicts for item in r["y"]]
        results["z[m]"] = [item for r in result_dicts for item in r["z"]]
        results["Offset[m]"] = [item for r in result_dicts for item in r["offset"]]
        results["CurrenName"] = [item for r in result_dicts for item in r["current"]]
        results["Return Period"] = [item for r in result_dicts for item in r["returnperiod"]]
        results["RAO"] = [item for r in result_dicts for item in r["rao"]]
        results["Draft"] = [item for r in result_dicts for item in r["initialz"]]
        
        

        # Selecting Max
        aux2 = []

        d = results["Return Period"].unique()
        list_Var = ["Sea surface Z", "Elevation"]
        for i, k in enumerate(d):
            name = []
            res = []
            for n, m in enumerate(varName):
                if m in [k for k in list_Var]:
                    name.append(m)
                    res.append(
                        results[results["Return Period"] == k].nlargest(1, ["{}".format(m)])
                    )
                else:
                    name.extend([m, m])
                    res.append(
                        results[results["Return Period"] == k].nlargest(1, ["{}".format(m)])
                    )
                    res.append(
                        results[results["Return Period"] == k].nsmallest(1, ["{}".format(m)])
                    )
            e = pd.concat(res)
            e["Criteria"] = name
            aux2.append(e)
        Max_Values = pd.concat(aux2)
        # Saving
        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                print("Printing VRA Results")
                # workBook.remove(workBook['VRA_Results'])
                # workBook.remove(workBook['VRA_Max'])
            except:
                print("Printing VRA Time Results")
            finally:
                print("Printing VRA Time Results")
                results.to_excel(writer, sheet_name="VRA_Time_Results")
                Max_Values.to_excel(writer, sheet_name="VRA_Time_Max")
        end = timer()

        print(
            f"Elapsed time - VRA_Time : {end - start}",
        )

    def CSYellowTail(self, Orcafile="file", Vessel_name="Vessel1"):

        """
        Creates current screening  (Far, Near,Cross and Transversal )
        Args:
            Orcafile:Orcaflex file with the line
        """
        start3 = timer()
        warnings.simplefilter("ignore")
        if os.path.isdir("01.CS") == False:
            os.mkdir("01.CS")

        if self.MetoceanRAO == True:
            aux_draft = self.data4["Draft"].unique()
        else:
            aux_draft = self.list_draft["Draft[m]"].unique()

        offAux = self.data4.iloc[:, 8:11]
        offAux.dropna(inplace=True)  # Offset data
        offAux.set_index(offAux.columns[0], inplace=True)

        riser_az = self.data3.iloc[1, 0]

        lst_current = self.data5  # List of current

        model = OrcFxAPI.Model(Orcafile)  # Orcaflexmodel

        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]
        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        LC = 0
        list_casename = []
        list_current = []
        list_OffSetDirData = []
        list_OffsetData = []
        list_Xoffset = []
        list_Yoffset = []
        list_initialz = []
        # list_tension=[]
        # list_bend=[]
        list_Offsetype = []
        list_rp = []
        list_bin = []
        ##

        ##
        list_par = [
            "Current_profile",
            "Current_dir[deg]",
            "Offset[m]",
            "Offset_dir[deg]",
            "X_offset[m]",
            "Y_offset[m]",
            "Draft[m]",
            # "Tension[kN]",
            # "Bend_radius[m]",
            "Offset_Type",
            "Return_period",
            "Sector",
        ]
        ##

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

        for m, n in tqdm(
            iterable=enumerate(setdir.index), desc="Offset-type", total=len(setdir.index)
        ):
            for i, j in tqdm(
                iterable=enumerate(lst_current.index),
                desc="Current-profile",
                total=len(lst_current.index),
            ):
                for k in aux_draft:
                    LC += 1
                    filename = "CS_{:04d}_{}".format(LC, j)
                    list_casename.append(filename)
                    environment = model.environment
                    environment.MultipleCurrentDataCanBeDefined = "Yes"
                    environment.SelectedCurrent = j
                    Auxoffdir = setdir[setdir.columns[0]][m]
                    list_OffSetDirData.append(Auxoffdir)
                    environment.RefCurrentDirection = Auxoffdir
                    environment.ActiveCurrent = j
                    list_current.append(j)
                    vessel = model[Vessel_name]
                    Off = offAux["Offset[m]"][lst_current["ReturnPeriod"][i]]
                    list_OffsetData.append(Off)
                    vessel.InitialZ = -k
                    list_initialz.append(-k)
                    vessel.InitialX = Off * np.cos(np.deg2rad(Auxoffdir))
                    list_Xoffset.append(vessel.InitialX)
                    vessel.InitialY = Off * np.sin(np.deg2rad(Auxoffdir))
                    list_Yoffset.append(vessel.InitialY)
                    model.SaveData(r"./01.CS/{}.dat".format(filename))
                    list_rp.append(lst_current["ReturnPeriod"][i])
                    list_bin.append(lst_current["Sector"][i])
                    list_Offsetype.append(n)
                    # model.CalculateStatics()
                    # line=model[line_name]
                    # bend=model[bend_name]
                    # aux_tension=line.StaticResult('Effective tension', OrcFxAPI.oeEndA)
                    # list_tension.append(aux_tension)
                    # aux_bend=bend.RangeGraph('Bend radius')
                    # bend_min=np.min(aux_bend.Mean)
                    # list_bend.append(bend_min)

        vet = zip(
            list_current,
            list_OffSetDirData,
            list_OffsetData,
            list_OffSetDirData,
            list_Xoffset,
            list_Yoffset,
            list_initialz,
            # list_tension,
            # list_bend,
            list_Offsetype,
            list_rp,
            list_bin,
        )
        result3 = pd.DataFrame(data=vet, index=list_casename, columns=list_par)

        # file='DataScreening_v2.xlsx'
        # self.ExcelFile

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing CS_Cases")
            except:
                print("Printing CS_Cases")
            finally:
                print("Printing CS_Cases")
                result3.to_excel(writer, sheet_name="CS_Cases_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )

    @staticmethod
    def _CSYellowTailThreadedRun(mnijklc, Orcafile, Vessel_name, offAux, lst_current):
        def _runModel(mnijklc, model, Vessel_name, offAux, lst_current, result_dict):
            ts = time.time()
            m, n, i, j, k, LC = (
                mnijklc[0],
                mnijklc[1],
                mnijklc[2],
                mnijklc[3],
                mnijklc[4],
                mnijklc[5],
            )
            filename = "CS_{:04d}_{}".format(LC, j)

            result_dict["casename"].append(filename)
            model.environment.MultipleCurrentDataCanBeDefined = "Yes"
            model.environment.SelectedCurrent = j
            Auxoffdir = m
            result_dict["OffSetDirData"].append(Auxoffdir)
            result_dict["OffSetDirData_curr"].append(Auxoffdir)
            model.environment.RefCurrentDirection = Auxoffdir
            model.environment.ActiveCurrent = j
            result_dict["current"].append(j)
            Off = offAux["Offset[m]"][lst_current["ReturnPeriod"][i]]
            result_dict["OffsetData"].append(Off)
            model[Vessel_name].InitialZ = -k
            result_dict["initialz"].append(-k)
            model[Vessel_name].InitialX = Off * np.cos(np.deg2rad(Auxoffdir))
            result_dict["Xoffset"].append(model[Vessel_name].InitialX)
            model[Vessel_name].InitialY = Off * np.sin(np.deg2rad(Auxoffdir))
            result_dict["Yoffset"].append(model[Vessel_name].InitialY)
            model.SaveData(r"./01.CS/{}.dat".format(filename))
            result_dict["rp"].append(lst_current["ReturnPeriod"][i])
            result_dict["bin"].append(lst_current["Sector"][i])
            result_dict["Offsetype"].append(n)
            te = time.time()
            # print(f"Internal run time: {te-ts}")
            return result_dict

        # ts = time.time()
        # result_dict = _runModel(mnijklc[0], Orcafile[0], Vessel_name[0], offAux[0], lst_current[0])
        # te = time.time()

        model = OrcFxAPI.Model(threadCount=3)  # Orcaflexmodel
        model.LoadDataMem(Orcafile[0])
        model.UseVirtualLogging()
        model.ForceInMemoryLogging()

        ts = time.time()
        # with ThreadingPool(len(mnijklc)) as tpool:
        #     result_lists = tpool.map(
        #         _runModel,
        #         mnijklc,
        #         Orcafile,
        #         Vessel_name,
        #         offAux,
        #         lst_current,
        #     )

        result_dict = defaultdict(list)
        for i, lc in enumerate(mnijklc):
            result_dict = _runModel(
                lc, model, Vessel_name[i], offAux[i], lst_current[i], result_dict
            )
        te = time.time()

        # print(f"Time: {te-ts}")
        return result_dict

    def ThreadedCSYellowTail(self, Orcafile="file", Vessel_name="Vessel1"):
        """
        Creates current screening  (Far, Near,Cross and Transversal )
        Args:
            Orcafile:Orcaflex file with the line
        """
        start3 = timer()
        warnings.simplefilter("ignore")
        if os.path.isdir("01.CS") == False:
            os.mkdir("01.CS")

        if self.MetoceanRAO == True:
            aux_draft = self.data4["Draft"].unique()
        else:
            aux_draft = self.list_draft["Draft[m]"].unique()
        lst_current = self.data5  # List of current
        list_offsetLabel = [
            "Near",
            "Far",
            "Cross-a",
            "Cross-b",
            "Cross-c",
            "Cross-d",
            "Trans-a",
            "Trans-b",
        ]

        with open(Orcafile, "rb") as f:
            Orcafile_bin = f.read()

        offAux = self.data4.iloc[:, 8:11]
        offAux.dropna(inplace=True)  # Offset data
        offAux.set_index(offAux.columns[0], inplace=True)

        list_offsetdir = [0, 180, 45, 135, 225, 315, 90, 270]
        list_offdir = []
        riser_az = self.data3.iloc[1, 0]

        for i in list_offsetdir:
            if i + riser_az > 360:
                list_offdir.append(i + riser_az - 360)
            else:
                list_offdir.append(i + riser_az)
        setdir = pd.DataFrame(index=list_offsetLabel)
        setdir["offdir"] = list_offdir

        mnijklc = []
        LC = 0
        for n in setdir.index:  # n out this loop
            for i, j in enumerate(lst_current.index):
                for k in aux_draft:
                    LC += 1
                    m = setdir.loc[n, "offdir"]
                    mnijklc.append((m, n, i, j, k, LC))

        threads_per_subprocess = int(len(mnijklc) / 6)
        mnijklc = [
            mnijklc[i : min([len(mnijklc), i + threads_per_subprocess])]
            for i in range(0, len(mnijklc), threads_per_subprocess)
        ]
        # self._CSYellowTailThreadedRun(
        #     mnijklc[0], [Orcafile_bin], [Vessel_name], [offAux], [lst_current]
        # )
        ts = time.time()
        result_dicts = []
        with ProcessingPool(len(mnijklc)) as pool:
            for p_result in tqdm(
                pool.imap(
                    self._CSYellowTailThreadedRun,
                    mnijklc,
                    [[Orcafile_bin] * len(mnijklc_i) for mnijklc_i in mnijklc],
                    [[Vessel_name] * len(mnijklc_i) for mnijklc_i in mnijklc],
                    [[offAux] * len(mnijklc_i) for mnijklc_i in mnijklc],
                    [[lst_current] * len(mnijklc_i) for mnijklc_i in mnijklc],
                )
            ):
                result_dicts.append(p_result)
        te = time.time()
        print(f"Total time: {te-ts}")
        list_par = [
            "Current_profile",
            "Current_dir[deg]",
            "Offset[m]",
            "Offset_dir[deg]",
            "X_offset[m]",
            "Y_offset[m]",
            "Draft[m]",
            # "Tension[kN]",
            # "Bend_radius[m]",
            "Offset_Type",
            "Return_period",
            "Sector",
        ]
        result_dict = {}
        for k in [
            "casename",
            "current",
            "OffSetDirData",
            "OffsetData",
            "OffSetDirData_curr",
            "Xoffset",
            "Yoffset",
            "initialz",
            "Offsetype",
            "rp",
            "bin",
        ]:
            result_dict[k] = [item for r in result_dicts for item in r[k]]

        result3 = pd.DataFrame().from_dict(result_dict)
        result3 = result3.set_index("casename")
        result3.columns = list_par

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing CS_Cases")
            except:
                print("Printing CS_Cases")
            finally:
                print("Printing CS_Cases")
                result3.to_excel(writer, sheet_name="CS_Cases_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Creating Cases  : {end3 - start3}",
        )

    def CS_Process(
        self,
        folder="./01.CS/",
        SheetName="CS_Cases_YellowTail",
        Line_name="Line1",
        Bend_name="Stiffener1",
    ):

        start3 = timer()
        model = OrcFxAPI.Model()
        list_tension = []
        list_bend = []
        list_bend_line=[]

        DATA = pd.read_excel(self.ExcelFile, sheet_name=SheetName, header=0, index_col=0)

        for i, j in tqdm(enumerate(DATA.index), desc="Cases", total=len(DATA.index)):
            file = folder + j + ".sim"
            model.LoadSimulation(file)
            line = model[Line_name]
            bend = model[Bend_name]
            aux_tension = line.StaticResult("Effective tension", OrcFxAPI.oeEndA)
            list_tension.append(aux_tension)
            aux_bend2=line.RangeGraph("Bend radius")
            bend_min2 = np.min(aux_bend2.Mean)
            list_bend_line.append(bend_min2)
            aux_bend = bend.RangeGraph("Bend radius")
            bend_min = np.min(aux_bend.Mean)
            list_bend.append(bend_min)

        DATA["Tension[kN]"] = list_tension
        DATA["Bend_radius[m]"] = list_bend
        DATA["Bend_radius_Line[m]"] = list_bend_line

        # Selecting Max
        aux2 = []
        d = DATA["Return_period"].unique()
        c = DATA["Sector"].unique()
        for i, k in enumerate(d):
            for n, m in enumerate(c):
                a = DATA[(DATA["Return_period"] == k) & (DATA["Sector"] == m)].nlargest(
                    1, ["Tension[kN]"]
                )
                a["Criteria"] = "Tension"

                b = DATA[(DATA["Return_period"] == k) & (DATA["Sector"] == m)].nsmallest(
                    1, ["Bend_radius[m]"]
                )
                b["Criteria"] = "Bending"
                f = DATA[(DATA["Return_period"] == k) & (DATA["Sector"] == m)].nsmallest(
                    1, ["Bend_radius_Line[m]"]
                )
                f["Criteria"] = "Bending_Line"
                # c=results[results['Return Period']==k].nlargest(1,['Pitch Amplitude [deg]'])
                # c['Criteria']='Pitch'
                # e=pd.concat([a,b,c])
                e = pd.concat([a, b,f])
                aux2.append(e)
        Max_Values = pd.concat(aux2)
        # Saving

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing CS_Cases")
            except:
                print("Printing CS_Cases")
            finally:
                print("Printing CS_Cases")
                DATA.to_excel(writer, sheet_name="CS_Results_YellowTail")
                Max_Values.to_excel(writer, sheet_name="CS_Max_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Processing Results  : {end3 - start3}",
        )

    @staticmethod
    def _CSProcessThreadedRun(j, folder, Line_name, Bend_name):
        def _runModel(j, folder, Line_name, Bend_name):
            ts = time.time()
            file = folder + j + ".sim"
            model = OrcFxAPI.Model()
            model.LoadSimulation(file)
            line = model[Line_name]
            bend = model[Bend_name]
            result_dict = defaultdict(list)
            aux_tension = line.StaticResult("Effective tension", OrcFxAPI.oeEndA)
            result_dict["tension"].append(aux_tension)
            aux_bend = bend.RangeGraph("Bend radius")
            bend_min = np.min(aux_bend.Mean)
            result_dict["bend"].append(bend_min)
            aux_bend2=line.RangeGraph("Bend radius")
            bend_min2 = np.min(aux_bend2.Mean)
            result_dict["bend_line"].append(bend_min2)
            te = time.time()
            print(f"Internal time: {te-ts}")
            return result_dict

        ts = time.time()
        with ThreadingPool(2) as tpool:
            result_lists = tpool.map(_runModel, j, folder, Line_name, Bend_name)
        te = time.time()
        print(f"ThreadPool Time: {te-ts}")
        return result_lists

    def ThreadedCS_Process(
        self,
        folder="./01.CS/",
        SheetName="CS_Cases_YellowTail",
        Line_name="Line1",
        Bend_name="Stiffener1",
    ):

        start3 = timer()

        DATA = pd.read_excel(self.ExcelFile, sheet_name=SheetName, header=0, index_col=0)

        j = []
        for di in DATA.index:
            j.append(di)

        threads_per_subprocess = 2
        j = [
            j[i : min([len(j), i + threads_per_subprocess])]
            for i in range(0, len(j), threads_per_subprocess)
        ]
        
        result_dicts=[]

        with ProcessingPool(5) as pool:
            for p_results in tqdm (
                pool.imap(
                    self._CSProcessThreadedRun,
                    j,
                    [[folder] * len(j_i) for j_i in j],
                    [[Line_name] * len(j_i) for j_i in j],
                    [[Bend_name] * len(j_i) for j_i in j],
                )
            ):
                result_dicts.append(p_results)
              

        DATA["Tension[kN]"] = [rr["tension"][0] for r in result_dicts for rr in r]
        DATA["Bend_radius[m]"] = [rr["bend"][0] for r in result_dicts for rr in r]
        DATA["Bend_radius_line[m]"] = [rr["bend_line"][0] for r in result_dicts for rr in r]

        # Selecting Max
        aux2 = []
        d = DATA["Return_period"].unique()
        c = DATA["Sector"].unique()
        for i, k in enumerate(d):
            for n, m in enumerate(c):
                a = DATA[(DATA["Return_period"] == k) & (DATA["Sector"] == m)].nlargest(
                    1, ["Tension[kN]"]
                )
                a["Criteria"] = "Tension"

                b = DATA[(DATA["Return_period"] == k) & (DATA["Sector"] == m)].nsmallest(
                    1, ["Bend_radius[m]"]
                )
                b["Criteria"] = "Bending"
                # c=results[results['Return Period']==k].nlargest(1,['Pitch Amplitude [deg]'])
                # c['Criteria']='Pitch'
                # e=pd.concat([a,b,c])
                f = DATA[(DATA["Return_period"] == k) & (DATA["Sector"] == m)].nsmallest(
                    1, ["Bend_radius_line[m]"]
                )
                f["Criteria"] = "Bending_Line"
                e = pd.concat([a, b, f])
                aux2.append(e)
        Max_Values = pd.concat(aux2)
        # Saving

        #test = 1

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing CS_Cases")
            except:
                print("Printing CS_Cases")
            finally:
                print("Printing CS_Cases")
                DATA.to_excel(writer, sheet_name="CS_Results_YellowTail")
                Max_Values.to_excel(writer, sheet_name="CS_Max_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Processing Results  : {end3 - start3}",
        )
    
    
    def Get_Max(
        self,
        SheetName="CS_Results_YellowTail",
    ):

        start3 = timer()

        DATA = pd.read_excel(self.ExcelFile, sheet_name=SheetName, header=0, index_col=0)
        
        # Selecting Max
        aux2 = []
        d = DATA["Return_period"].unique()
        c = DATA["Current_profile"].unique()
        for i, k in enumerate(d):
            for n, m in enumerate(c):
                a = DATA[(DATA["Return_period"] == k) & (DATA["Current_profile"] == m)].nlargest(
                    1, ["Tension[kN]"]
                )
                a["Criteria"] = "Tension"

                b = DATA[(DATA["Return_period"] == k) & (DATA["Current_profile"] == m)].nsmallest(
                    1, ["Bend_radius[m]"]
                )
                b["Criteria"] = "Bending"
                # c=results[results['Return Period']==k].nlargest(1,['Pitch Amplitude [deg]'])
                # c['Criteria']='Pitch'
                # e=pd.concat([a,b,c])
                f = DATA[(DATA["Return_period"] == k) & (DATA["Current_profile"] == m)].nsmallest(
                    1, ["Bend_radius_line[m]"]
                )
                f["Criteria"] = "Bending_Line"
                e = pd.concat([a, b, f])
                aux2.append(e)
        Max_Values = pd.concat(aux2)
        # Saving

        #test = 1

        with pd.ExcelWriter(
            self.ExcelFile, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # workBook = writer.book
            try:
                # workBook.remove(workBook['GA_Cases'])
                print("Printing CS_Cases")
            except:
                print("Printing CS_Cases")
            finally:
                print("Printing CS_Cases")
                Max_Values.to_excel(writer, sheet_name="CS_Graph_YellowTail")
        end3 = timer()
        print(
            f"Elapsed time-Processing Results  : {end3 - start3}",
        )


# Run
if __name__ == "__main__":
    try:
        Data = Screening(
            Vessel_name="Vessel",
            orcafile="BaseVessel_Extreme.dat",
            ExcelFile="./umb5/Seed1.xlsx",
            Save=False,
            CreateCases=True,
            MetoceanRAO=True,
            Acc=True,
            Spectrum=0,
        )

        Data.LoadCasesVRA()
        Data.RunVRAGA()
        Data.RunVRAGE("VRA_Max")
        Data.CreateWaveCasesAkerYellowTail("GE_Results", "VRA_Max")
        Data.CreateCurrentCasesAkerYellowTail("GE_Results", "VRA_Max")
        #Data.CreateCasesSheetGA("GE_Results", "VRA_Max")
        #Data.CreateWaveCasesKizomba("GE_Results", "VRA_Max")
        #Data.CreateCurrentCasesKizomba("GE_Results", "VRA_Max")

        # Data.VRATime(savefiles=True)
        # Data.ProcessVRATime("01.VRA_Time2")
        # Data.ThreadedCSYellowTail(
        #     Orcafile="BaseCaseYT_9in_PR_Rev1_Staggered.dat", Vessel_name="Vessel1"
        # )
        # Data.ThreadedCS_Process(
        #     folder="./01.CS/",
        #     SheetName="CS_Cases_YellowTail_2",
        #     Line_name="Line1",
        #     Bend_name="Stiffener1",
        # )
        
        # Data.Get_Max(
        #     SheetName="CS_Results_YellowTail",
        # )

    except PermissionError:
        print("ERROR: Please, close the excel file")
    except OrcFxAPI.DLLError:
        print("ERROR: Please, check Orcaflex license")
