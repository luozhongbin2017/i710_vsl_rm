import pandas as pd
import numpy as np
import os, re, math
from datetime import datetime

def findSourceType(VehType, M):
    '''
    Map VISSIM vehicle type into MOVES source type
    VehType: VISSIM vehicle type number
    M: vehicle weight in metric tons
    '''
    if 100 == VehType:
        return 21
    if 200 ==VehType:
        if 15 < M <= 25:
            return 41
        if 25 < M:
            return 62
        if 3.8 < M < 6.3:
            return 53
        if 1.5 < M <= 3.8:
            return 32
    if 300 == VehType:
        return 42
    else:
        return 21


def findParameters(sourceType):
    '''Find vehicle parameters for specific source type'''
    if 21 == sourceType:
        A = 0.156461
        B = 0.002002
        C = 0.000493
        M = 1.4788
        f = 1.4788
    elif 31 == sourceType:
        A = 0.22112
        B = 0.002838
        C = 0.000698
        M = 1.86686
        f = 1.86686
    elif 32 == sourceType:
        A = 0.235008
        B = 0.003039
        C = 0.000748
        M = 2.05979
        f = 2.05979        
    elif 41 == sourceType:
        A = 1.29515
        B = 0        
        C = 0.003715
        M = 19.5937
        f = 17.1
    elif 42 == sourceType:
        A = 1.0944
        B = 0
        C = 0.003587
        M = 16.556
        f = 17.1
    elif 52 == sourceType:
        A = 0.561933
        B = 0
        C = 0.001603
        M = 7.64159
        f = 17.1
    elif 53 == sourceType:
        A = 0.498699
        B = 0
        C = 0.001474
        M = 6.25047
        f = 17.1    
    elif 61 == sourceType:
        A = 1.96354
        B = 0
        C = 0.004031
        M = 29.3275
        f = 17.1
    else:
        A = 2.08126
        B = 0
        C = 0.004188
        M = 31.4038
        f = 17.1
    return A, B, C, M, f

def findVSP(A, B, C, M, f, v, a, grad):
    '''
    Find VSP of vehicle for current time
    '''
    tan_theta = grad/100
    sin_theta = math.sqrt(tan_theta**2 / (1 + tan_theta**2))
    VSP = (A*v + B*(v**2) + C*(v**3) + M*(a + 9.8*sin_theta)*v) / f
    return VSP


def findOperationMode(VSP, v, a):
    '''
    Find operating mode of vehicle for current time
    The operating mode number returned by this function is not MOVES operation mode ID, but the link ID 1~23, corresponding to 23 operation modes, instead.
    This is for convenience in table lookup
    '''
    if a <= -2:
        return 1
    elif v < 1 and v > -1:
        return 2
    elif v < 25:
        if VSP < 0: return 3
        if VSP < 3: return 4
        if VSP < 6: return 5
        if VSP < 9: return 6
        if VSP < 12: return 7
        return 8
    elif v < 50:
        if VSP < 0: return 9
        if VSP < 3: return 10
        if VSP < 6: return 11
        if VSP < 9: return 12
        if VSP < 12: return 13
        if VSP < 18: return 14
        if VSP < 24: return 15
        if VSP < 30: return 16
        return 17
    else:
        if VSP < 6: return 18
        if VSP < 12: return  19
        if VSP < 18: return 20
        if VSP < 24: return 21
        if VSP < 30: return 22
        return 23



def fzpEvaluation(fzpPath, startTime, stopTime):
    '''Evaluate the fzp file of VISSIM
    fzpPath: the path of *.fzp file'''

    mph2ms = 0.44704    #convert mph to m/s
    fts2ms = 0.3048 #convert ft/s^2 to m/s^2
    vehData = pd.read_csv(fzpPath, skiprows = 21, sep = ';', skipinitialspace = True)

    # abandon strings in the last column
    vehData = vehData[vehData.t <= stopTime]
    vehData = vehData[list(vehData.columns[:-1])]
    vehData.dropna(inplace = True)
    vehData.sort(['VehNr', 't'], inplace = True)
    nVeh = int(max(vehData.VehNr))
    throughput = 0
    VMT = 0
    travelTime = 0
    nStop = 0
    nLC = 0
    HC = 0
    CO = 0
    NOx = 0
    CO2 = 0
    energy = 0
    PM25 = 0

    emissionRates = pd.read_csv('lookuptable.csv')


    for i in xrange(1, nVeh+1):
        veh = vehData[i == vehData.VehNr]
        # vehData_temp = vehData[i != vehData.VehNr]
        # del vehData
        # vehData = vehData_temp
        # del vehData_temp
        if 0 == len(veh):
            continue
        if veh.iloc[-1]['t'] < startTime or veh.iloc[-1]['t'] > stopTime:
            continue
        if int(veh.iloc[-1]['Link']) not in [304, 306]:
            continue

        throughput += 1
        #print throughput
        travelTime += (veh.iloc[-1]['t'] - veh.iloc[0]['t']) / 60
        nStop += veh.iloc[-1]['Stops']
        #lc = veh[veh.LCh != '-']
        # lc = list(veh.LCh)
        # for j in xrange(len(lc)-1):
        #     if lc[j] != lc[j+1]:
        #         nLC += 0.5
        # speed = list(veh.v)
        # for sp in speed:
        #     VMT += sp / 

        vehType = veh.Type.iloc[0] 
        weight = veh.Weight.iloc[0]
        sourceType  = findSourceType(vehType, weight)
        A, B, C, M, f = findParameters(sourceType)
        emissions = emissionRates[emissionRates.sourceTypeID == sourceType]
        sp = list(veh.v)
        acc = list(veh.a)
        lc = list(veh.LCh)

        for sec in range(len(veh) - 1):
            # v = veh.iloc[sec]['v'] 
            # a = veh.iloc[sec]['a']
            v = sp[sec]
            a = acc[sec]
            # grad = veh.iloc[sec]['Grad']
            grad = 0
            VSP = findVSP(A, B, C, M, f, v * mph2ms, a * fts2ms, grad)
            operationMode = findOperationMode(VSP, v, a)
            # emission = emissions[emissions.linkID == operationMode]
            # VMT += v / 3600
            # HC += float(emission[emission.pollutantID == 1]['ratePerDistance']) / 3600
            # CO += float(emission[emission.pollutantID == 2]['ratePerDistance']) / 3600
            # NOx += float(emission[emission.pollutantID == 3]['ratePerDistance']) / 3600
            # CO2 += float(emission[emission.pollutantID == 90]['ratePerDistance']) / 3600
            # energy += float(emission[emission.pollutantID == 91]['ratePerDistance']) / 3600
            # PM25 += float(emission[emission.pollutantID == 110]['ratePerDistance']) / 3600
            emission = list(emissions[emissions.linkID == operationMode]['ratePerDistance'])
            VMT += v / 3600
            HC += emission[0] / 3600
            CO += emission[1] / 3600
            NOx += emission[2] / 3600
            CO2 += emission[3] / 3600
            energy += emission[4] / 3600
            PM25 += emission[5] / 3600

            if lc[sec] != lc[sec+1]:
                nLC += 0.5




    aTravelTime = travelTime / throughput
    aStops = float(nStop) / throughput
    aLC = nLC / throughput
    aHC = HC / VMT
    aCO = CO / VMT
    aNOx = NOx / VMT
    aCO2 = CO2 / VMT
    aEnergy = energy / VMT
    aPM25 = PM25 / VMT

    return throughput, aTravelTime, VMT, aStops, aLC, aHC, aCO, aNOx, aCO2, aEnergy, aPM25


def eval_main():
    inc = [1]
    ctrl = [(0, 0), (1, 2)]
    MC = [3*i+1 for i in range(5)]
    # inc = [10]
    # ctrl = [(1, 2)]
    # MC = [1, 4]
    incTime = [(1200, 2400), (1200, 3600), (1500, 3900)]
    for i in inc:
        startTime, stopTime = incTime[i]
        for j in ctrl:
            folderDir = './incident%dmin/Ctrl%d/LC%d' % (i, j[0], j[1])
            resultFile = open(os.path.join(folderDir, 'record.txt'), 'w')
            average = []
            for k in MC:                
                fzpPath = folderDir + '/I710_%d.fzp' %k
                throughput, aTravelTime, VMT, aStops, aLC, aHC, aCO, aNOx, aCO2, aEnergy, aPM25 = fzpEvaluation(fzpPath, startTime, stopTime)
                resultFile.write(str([throughput, aTravelTime, VMT, aStops, aLC, aHC, aCO, aNOx, aCO2, aEnergy, aPM25]) + '\n')
                average.append([throughput, aTravelTime, VMT, aStops, aLC, aHC, aCO, aNOx, aCO2, aEnergy, aPM25])
            average = np.array(average).mean(axis = 0)                
            resultFile.write(str(average) + '\n')
            resultFile.close()
                # print fzpPath

if __name__ == '__main__':
    eval_main()
