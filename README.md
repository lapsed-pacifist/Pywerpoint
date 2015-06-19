# Pywerpoint
Easy to use tools for import data into powerpoint

Allows array views of win32com.client tables instead of horrible vba style interfaces.

```
T = pd.DataFrame(np.arange(45).reshape(9,5), columns=list('ABCDE'))
with Presentation(template='table_temp.pptx') as P:
	table = P[0].tables[0] #list tables on slides
	print table.cells.text
	table.cells.text = T
	
	P[1] = 'Blank' # insert slide with a powerpoint defined layout
	P[2] = 'table1' # insert slide by specifying the title of a template slide
	
	P[1].create_table(2, 2) # insert table
	
	images = P[0].images		
```
