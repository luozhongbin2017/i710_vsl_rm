import os
result_file = open('evalResult.txt', 'r')
result = [0.0] * 4
for line in result_file:
	#print line
	item = line.split()
	#print item
	for i in range(len(result)):
		result[i] += float(item[i])
for i in range(len(result)):
	result[i] /= 10
	result[i] *= 2500
	print int(result[i])
result_file.close()

