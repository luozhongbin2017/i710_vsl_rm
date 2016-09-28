import pandas as pd
import numpy as np
import os, re
from datetime import datetime

def read_sum(sumFileName):
	sumFile = open(sumFileName, 'r')
	flag = 0
	result = {}
	for line in sumFile:
		if 0 == flag:
			match = re.search('Distance Traveled', line)
			if match:
				flag = 'dist'
				continue
			match = re.search('Fuel Use', line)
			if match:
				flag = 'fuel'
				continue
			match = re.search('Tailpipe Out Emissions', line)
			if match:
				flag = 'emission'
				continue
		elif 'dist' == flag:
			value = line.split()
			dist = float(value[0])
			flag = 0
		elif 'fuel' == flag:
			value = line.split()
			result[flag] = float(value[0])
			flag = 0
		elif 'emission' == flag:
			value = line.split()
			if 3 < len(value):
				result[value[0]] = float(value[2])
			else:
				flag = 0
	sumFile.close()
	for key in result.keys():
		result[key] *= dist
	return dist, result.values()

#result = read_sum('i710-sum')




		





def fzpEvaluation():
	'''Evaluate the fzp file of VISSIM
	fzpPath: the path of *.fzp file'''
	data = pd.read_csv('i710.fzp', skiprows = 18, sep = ';', skipinitialspace = True)
	data = data[list(data.columns[:-1])]
	data.dropna(inplace = True)
	#data.reindex()
	data.sort(['VehNr', 't'], inplace = True)
	nVeh = int(max(data.VehNr))
	carThroughput = 0
	truckThroughput = 0
	carDistance = 0
	truckDistance = 0
	carRecord = []
	truckRecord = []
	throughput = 0
	#for i in range(1,2):
	for i in range(1,nVeh+1):
		#print i
		veh = data[i == data.VehNr]
		if 0 == len(veh):
			continue

		#print veh
		if veh.iloc[-1]['t'] >= 600:
			if int(veh.iloc[-1]['Link']) in [304, 306]:
				throughput += 1
				#print throughput, i
				travelTime = (veh.iloc[-1]['t'] - veh.iloc[0]['t']) / 60
				noStop = veh.iloc[-1]['Stops']
				#lc =  veh[veh.LCh != '-']
				lc = list(veh.LCh)
				noLCh = 0
				for j in xrange(len(lc)-1):
					if lc[j] != lc[j+1]:
						noLCh += 1
				# for j in xrange(len(lc)-1):
				# 	if lc.iloc[j]['LCh'] != lc.iloc[j+1]['LCh']:
				# 		noLCh += 1
				noLCh /= 2


				# for j in xrange(len(veh)-1):
				# 	if veh.iloc[j]['Lane'] != veh.iloc[j+1]['Lane'] and veh.iloc[j]['Link'] == veh.iloc[j+1]['Link']:
				# 		noLCh += 1


				#noLCh = len(veh['LCh'][veh['LCh'] != '-'])
				#vehRecord.append([travelTime, noStop, noLCh])
				actFile = open('i710-act', 'w')
				ctrFile = open('i710-ctr', 'w')
				for sec in range(len(veh)) :
					#pass
					#actFile.write(str(sec+1) + ',' + str(veh.iloc[sec]['v']) + ',' + str(veh.iloc[sec]['a'] * 3600 / 5280) + '\n')
					actFile.write(str(sec+1) + ',' + str(veh.iloc[sec]['v']) + '\n')
				actFile.close()
				
				if 100 == veh.iloc[0]['Type']:
					carThroughput += 1
					ctrFile.write('IN_UNITS        = ENGLISH\n\n')
					ctrFile.write('OUT_UNITS        = ENGLISH\n\n')
					ctrFile.write('VEHICLE_CATEGORY = 5\n\n')
					ctrFile.write('CO2_OUT = ON')
					ctrFile.close()
					os.system('cmemCore i710')
					outputFileName = 'i710-sum'
					[dist, environment] = read_sum(outputFileName)
					carRecord.append([travelTime, noStop, noLCh] + environment)
					carDistance += dist

				elif 200 == veh.iloc[0]['Type']:
					ctrFile.write('IN_UNITS        = ENGLISH\n\n')
					ctrFile.write('OUT_UNITS        = ENGLISH\n\n')
					ctrFile.write('VEHICLE_CATEGORY = 7\n\n')
					ctrFile.write('massTrailer_lb = ' + str(int(veh.iloc[0]['Weight'] * 2204)))
					ctrFile.close()
					os.system('hddCore i710-ctr i710-act')
					outputFileName = 'i710.sum'
					[dist, environment] = read_sum(outputFileName)
					truckRecord.append([travelTime, noStop, noLCh] + environment)
					truckThroughput += 1
					truckDistance += dist
				
				# if 100 == veh.iloc[0]['Type']:
				# 	carThroughput += 1
				# 	carRecord.append([travelTime, noStop, noLCh])
				# elif 200 == veh.iloc[0]['Type']:
				# 	truckThroughput += 1
				# 	truckRecord.append([travelTime, noStop, noLCh])

	print throughput
	carRecord = np.array(carRecord)
	truckRecord = np.array(truckRecord)
	carTotal = list(carRecord.sum(axis = 0))
	truckTotal = list(truckRecord.sum(axis = 0))
	for j in range(3, len(carTotal)):
		carTotal[j] /= carDistance
		truckTotal[j] /= truckDistance

	pass
	# for idx in range(-5,0):
	# 	carTotal[idx] /= carDistance
	# 	truckTotal[idx] /= truckDistance

	return carTotal, truckTotal

# ts = datetime.now().time()
# [carTotal, truckTotal] = fzpEvaluation()
# te = datetime.now().time()
# print ts, '\n'
# print te, '\n'
# print carTotal, '\n'
# print truckTotal, '\n'
# pass

