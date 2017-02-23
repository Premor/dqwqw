import sys
import xlrd
import json
from tkinter import *
from tkinter.filedialog import *
 
def do_some():
	path = askopenfilename()
	rb = xlrd.open_workbook(path)#,formatting_info=True)
	sheet = rb.sheet_by_index(0)
	print(str(sheet.nrows)+' '+str(sheet.ncols))
	d={}
	z=[]
	f=False
	x=0
	for colx in range(sheet.ncols):
		if not f:
			x=colx
		for rowx in range(sheet.nrows):
			if f:
				d.update({sheet.cell_value(rowx,x):sheet.cell_value(rowx,colx)})
			else:
				d.update({sheet.cell_value(rowx,colx):''})
		if f:
			buf=d.copy()
			z.append(buf)
			print (z)
		f=True
	b=json.dumps(z)
	print(z)
	f=open('tmp.txt','w')
	f.write(b)
	f.close()



#path=sys.executable
path=''
root = Tk()
root.geometry("800x600+200+200")
exit=Button(root,text='exit',command=lambda:sys.exit())
do_something=Button(root,text='do',command=do_some)
exit.pack()
do_something.pack()
 
root.mainloop()
