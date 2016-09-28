import pandas as pd
import numpy as np
import os

nMC = 5
randomSeeds = range(nMC)
for i in range(nMC):
	randomSeeds[i] = randomSeeds[i]*3 + 1

dataMC = np.array([0, 0, 0])
for randomSeed in randomSeeds:
	bottleneckFileName = 'bottleneckData_%d.txt' % randomSeed
	# print bottleneckFileName
	# data = pd.read_csv(bottleneckFileName, skiprows = 118, sep = '\t', header = 0, names = ['t', 'occ', 'v'])
	data = np.loadtxt(bottleneckFileName, skiprows = 119)
	dataMean = data.mean(axis = 0)
	dataMC = np.vstack([dataMC, dataMean])
	#print dataMC
	#print data
	#os.system('pause')
dataMC = dataMC[1:, :]
print dataMC
finalMean = dataMC.mean(axis = 0)
print finalMean

	#bottleneckFile = open()
