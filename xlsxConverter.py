import xlsxwriter
import os


def realOpen(fileName):
	return(open(os.path.join(os.getcwd(), fileName)))#uses true path to open file

os.mkdir(os.path.join(os.getcwd(), "Results"))

summarybook = xlsxwriter.Workbook("Summary.xlsx")#summarry spreadsheet
summary = summarybook.add_worksheet()
#assert(False)
summary.write("B1", "X Max")
summary.write("C1", "X Min")
summary.write("D1", "Y Max")
summary.write("E1", "Y Min")#summarry setup


dirCounter=0#provides starting number for nameing convention. should be changed to whatever the lowest number is i.e if lowest is ALL0021, this should be set to 21
#begin loop
absolute = 0#absolute number of loops, regardless of dirCounter start
while(1):
	if dirCounter>=1000:
		currentDirNum = str(dirCounter)
	elif dirCounter>=100:
		currentDirNum = '0'+str(dirCounter)
	elif dirCounter>=10:
		currentDirNum = '00'+str(dirCounter)
	else:
		currentDirNum = '000'+str(dirCounter)

	print os.path.join(os.getcwd(),  "ALL"+currentDirNum+"/")
	f1Data = []
	f2Data = []
	f3Data = []
	f4Data = []
	fileList = []
	dataList = [f1Data,f2Data,f3Data,f4Data]
	countList = [0,0,0,0]#stores value for number of lines each file has, (1,2,3,4)
	
	try:
		f1 = realOpen("ALL"+currentDirNum+"/F"+currentDirNum+"CH1.CSV")
		
	except Exception:
		f1 = None
		print("1st file not found, terminating")
		break
	try:
		f2 = realOpen("ALL"+currentDirNum+"/F"+currentDirNum+"CH2.CSV")

	except Exception:
		f2 = None
		print("2nd file not found, proceeding")
	try:
		f3 = realOpen("ALL"+currentDirNum+"/F"+currentDirNum+"CH3.CSV")
		
	except Exception:
		f3 = None
		print("3rd file not found, proceeding")
	try:
		f4 = realOpen("ALL"+currentDirNum+"/F"+currentDirNum+"CH4.CSV")
	except Exception:
		f4 = None
		print("4th file not found, proceeding")
		
	fileList.append(f1)
	fileList.append(f2)
	fileList.append(f3)
	fileList.append(f4)
	
	for i in range(len(fileList)):#only perform rips for confirmed loaded files, in order.
		while(1):#rip file into string
			countList[i]=countList[i]+1
			#rawList = fileList[i].readline().split(',')
			#print rawList
			#print rawList[4]
			try:
				dataList[i].append((fileList[i].readline().split(','))[4].strip())
			except IndexError:
				fileList[i].close()
				break
			except AttributeError:
				break
	
	addListX=[]#sum of x channels for text files
	addListY=[]#sum of y channels
	if(dataList[1]):
		for i in range(len(dataList[0])-1):
			addListX.append(float(dataList[0][i])+float(dataList[1][i]))
	elif(dataList[0]):
		addListX=dataList[0]
	else:
		addListX = [0]
	
	if(dataList[3]):
		for i in range(len(dataList[2])-1):
			addListY.append(float(dataList[2][i])+float(dataList[3][i]))
	elif(dataList[2]):
		print "hel"
		addListY=dataList[2]
	else:
		print "goby"
		addListY = [0]
		print addListY
		
	print addListY
	
	summary.write("A"+str(absolute+2), "F"+currentDirNum)
	summary.write("B"+str(absolute+2), str(max(addListX)))
	summary.write("C"+str(absolute+2), str(min(addListX)))
	summary.write("D"+str(absolute+2), str(max(addListY)))
	summary.write("E"+str(absolute+2), str(min(addListY)))
	
	workbook = xlsxwriter.Workbook('F'+currentDirNum+".xlsx")
	worksheet = workbook.add_worksheet()
	#ws settup
	worksheet.write('D1', 'time')
	worksheet.write('E1', 'X1')
	worksheet.write('F1', 'X2')
	worksheet.write('G1', 'Y1')
	worksheet.write('H1', 'Y2')
	worksheet.write('I1', 'X Sum')
	worksheet.write('J1', 'Y Sum')
	worksheet.write('L2', 'Xmax')
	worksheet.write('L3', 'Xmin')
	worksheet.write('L4', 'X average')
	worksheet.write('L5', 'Ymax')
	worksheet.write('L6', 'Ymin')
	worksheet.write('L7', 'Y average')
	#add data
	j = 2
	for i in f1Data:
		worksheet.write('E'+str(j), i)
		worksheet.write_formula('I'+str(j), "=E"+str(j)+"+F"+str(j))#write xsum
		j+=1 
	j = 2
	for i in f2Data:	
		worksheet.write('F'+str(j), i)
		j+=1 
	j = 2
	for i in f3Data:
		worksheet.write('G'+str(j), i)
		worksheet.write_formula('J'+str(j), "=G"+str(j)+"+H"+str(j))#write ysum
		j+=1 
	j = 2
	for i in f4Data:	
		worksheet.write('H'+str(j), i)
		j+=1 
	
	#add in calc
	worksheet.write_formula("M2", "=MAX(I2:I"+str(countList[0])+")")
	worksheet.write_formula("M3", "=MIN(I2:I"+str(countList[0])+")")
	worksheet.write_formula("M4", "=AVERAGE(I2:I"+str(countList[0])+")")
	worksheet.write_formula("M5", "=MAX(J2:J"+str(countList[2])+")")
	worksheet.write_formula("M6", "=MIN(J2:J"+str(countList[2])+")")
	worksheet.write_formula("M7", "=AVERAGE(J2:J"+str(countList[2])+")")
	
	workbook.close()
	os.rename(os.path.join(os.getcwd(), 'F'+currentDirNum+".xlsx"), os.path.join(os.getcwd(), "Results/F"+currentDirNum+".xlsx"))#move newly made spreadsheet into reults folder
	dirCounter+=1
	absolute+=1
summarybook.close()
os.rename(os.path.join(os.getcwd(), "Summary.xlsx"), os.path.join(os.getcwd(), "Results/Summary.xlsx"))#move summarry sheet into results folder
