import os
import pandas as pd
import matplotlib.pyplot as plt

def plotDensityCurve(DataDir, randomSeed, sec):
    densityDir = os.path.join(DataDir, "densityLog_%d.txt"%randomSeed)
    densityData = pd.read_csv(densityDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    densityData = densityData[list(densityData.columns[:-1])]
    #print densityData
    t = densityData['t/s']
    den = densityData['Sec%d'%sec]
    den = den.rolling(window = 10, center = False).mean()
    plt.plot(t, den, linewidth = 2)
    plt.grid()
    plt.show()

def plotDensityContour(DataDir, randomSeed):
    densityDir = os.path.join(DataDir, "densityLog_%d.txt"%randomSeed)
    densityData = pd.read_csv(densityDir, skiprows = 1, sep = '\t', skipinitialspace = True)
    densityData = densityData[list(densityData.columns[:-1])]




if __name__ == '__main__':
    DataDir = "./data/scenario0/ctrl0/LC2"
    randomSeed = 1
    sec = 9
    plotDensityCurve(DataDir, randomSeed, sec)
