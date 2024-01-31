from openpyxl import load_workbook



#1 lê pasta de trablho e planilha 

Wb=load_workbook("data/pivot_table.xlsx")

sheet = Wb["relatorio"]

# 2 acessando o valor especifico

#print(sheet["a3"].value)
#print(sheet["b3"].value)

###############


# interando valores com loop

for i in range(2, 6):
    ano = sheet["A%s" %i].value 
    am = sheet["b%s" %i].value 
    bt = sheet["c%s" %i].value 
    
    print("{0} o aston maintin vendeu {1} e o bentley vendeu {2}".format( ano , am , bt))
    
  