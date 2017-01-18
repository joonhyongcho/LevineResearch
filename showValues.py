import xlwings as xw

def getRowSums(values):

	sums = [0 for i in range(len(values))]

	for row in range(len(values)):
		for col in range(len(values[row])):
			if values[row][col]:
				sums[row] += values[row][col] 

	return sums

def getColSums(values):

	sums = [0 for i in range(len(values[0]))]

	for row in values:
		for idx in range(len(row)):
			if row[idx]:
				sums[idx] += row[idx]

	return sums

def divideByRowAndColSums(values):

	rowSums = getRowSums(values)

	# divide each col in each row by row sum
	for row in range(len(values)):
		for col in range(len(values[row])):
			if values[row][col]:
				values[row][col] /= rowSums[row]

	print(values)
	colSums = getColSums(values)

	# divide each col by colsum 
	for row in range(len(values)):
		for col in range(len(values[row])):
			if values[row][col]:			
				values[row][col] /= colSums[col]

	return values

def divideThroughXTimes(values, x):

	for i in range(x):
		values = divideByRowAndColSums(values)

	return values

wb = xw.Book('finance.xlsx') 

values = wb.sheets[1].range('B2:H39').value

print(divideThroughXTimes(values, 2))



