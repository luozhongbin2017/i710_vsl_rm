import os
import pandas as pd
import matplotlib.pyplot as plt

def plotDensityCurve(DataDir, randomSeed, sec):
    densityDir = os.path.join(DataDir, "densityLog_%d.txt"%randomSeed)
    densityData = pd.read_csv(densityDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    #densityData = densityData[list(densityData.columns[:-1])]
    secLen = [1494.979, 2268.889]
    densityData['Sec17'] = (densityData['Sec17'] * secLen[0] + densityData['Sec18'] * secLen[1]) / sum(secLen)
    densityData = densityData[list(densityData.columns[:-2])]
    t = densityData['t/s']
    den = densityData['Sec%d'%sec]
    den = den.rolling(window = 10, center = False).mean()
    plt.figure()
    plt.plot(t, den, linewidth = 2)
    plt.grid()
    #plt.show()

def plotDensityContour(DataDir, randomSeed):
    densityDir = os.path.join(DataDir, "densityLog_%d.txt"%randomSeed)
    densityData = pd.read_csv(densityDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    #densityData = densityData[list(densityData.columns[:-1])]
    secLen = [1494.979, 2268.889]
    densityData['Sec17'] = (densityData['Sec17'] * secLen[0] + densityData['Sec18'] * secLen[1]) / sum(secLen)
    densityData = densityData[list(densityData.columns[:-2])]
    X = range(1,9)
    Y = densityData['t/s']
    Z = densityData[["Sec%d"%num for num in range(10, 18)]]
    plt.figure()
    plt.contourf(X, Y, Z)
    plt.grid()
    #plt.show()

def plotFlowCurve(DataDir, randomSeed, sec):
    flowDir = os.path.join(DataDir, "flowSectionLog_%d.txt"%randomSeed)
    flowData = pd.read_csv(flowDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    flowData = flowData[list(flowData.columns[:-1])]    
    t = flowData['t/s']
    fl = flowData['Sec%d'%sec]
    fl = fl.rolling(window = 10, center = False).mean()
    plt.figure()
    plt.plot(t, fl, linewidth = 2)
    plt.grid()

def plotVSL(DataDir, randomSeed, sec):
    fileDir = os.path.join(DataDir, "vslLog_%d.txt"%randomSeed)
    Data = pd.read_csv(fileDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    t = Data['t/s']
    Data = Data[['Sec%d'%num for num in range(10, 17)]]
    plt.figure()
    plt.step(t, Data['Sec%d'%sec])

def plotQue(DataDir, randomSeed, rampID):
    rampList = [64, 61, 54, 49]
    fileDir = os.path.join(DataDir, 'rampQueLog_%d'%randomSeed)
    Data = pd.read_csv(fileDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    t = Data['t/s']
    Data = Data[['Ramp%d'%num for num in rampList]]
    plt.figure()
    plt.step(t, Data['Ramp%d'%rampID])



if __name__ == '__main__':
    DataDir = "./data/scenario0/ctrl0/LC2"
    randomSeed = 1
    sec = 10
    plotDensityCurve(DataDir, randomSeed, sec)
    # plotDensityContour(DataDir, randomSeed)
    # plotFlowCurve(DataDir, randomSeed, sec)
    # plotVSL(DataDir, randomSeed, sec)
    plt.show()
