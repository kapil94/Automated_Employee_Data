import openpyxl,sqlite3



def addRecords():
	
	Emp_id=checkRecords()
	
	wb=openpyxl.load_workbook('Employee.xlsx')
	sheet=wb.active.title
	conn=sqlite3.connect('Employee.db')
	
	if Emp_id!=0 and Emp_id!=wb[sheet].cell(row=wb[sheet].max_row,column=1).value:
		
		for row in range(2,wb[sheet].max_row+1):
			
			if Emp_id==wb[sheet].cell(row=row,column=1).value and row!=wb[sheet].max_row:
				new_row=row+1
				
				for row in range(new_row,wb[sheet].max_row+1):
					
					sql='INSERT INTO Employee values ({0},"{1}",{2})'.format(wb[sheet].cell(row=row,column=1).value,wb[sheet].cell(row=row,column=2).value,wb[sheet].cell(row=row,column=3).value)
				
					conn.execute(sql)
					conn.commit()
	
	
	elif Emp_id==wb[sheet].cell(row=wb[sheet].max_row,column=1).value:
		print('Records already present!!')
								
	else:
		
		for row in range(2,wb[sheet].max_row+1):
			
			sql='INSERT INTO Employee values ({0},"{1}",{2})'.format(wb[sheet].cell(row=row,column=1).value,wb[sheet].cell(row=row,column=2).value,wb[sheet].cell(row=row,column=3).value)
				
			conn.execute(sql)
			conn.commit()
		
		
def checkRecords():
	Emp_id=0
	conn=sqlite3.connect('Employee.db')
	
	sql='select Emp_id from Employee ORDER BY Emp_id DESC LIMIT 1'
	result=conn.execute(sql)
	for emp_id in result:
		Emp_id=emp_id[0]
	
	return Emp_id


addRecords()
