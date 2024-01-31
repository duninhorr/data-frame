import pandas as pd 

data = pd.read_excel("data/VendaCarros.Xlsx")

print(type(data))

df=data[["Fabricante","ValorVenda","Ano"]]
print(df)


#criando tabela divo ( tabela dinamica )

pivot_table = df.pivot_table( 
 index="Ano", 
 columns="Fabricante",
 values= "ValorVenda",                             
 aggfunc= "sum"                           
)

print(pivot_table)

#exportando tabela pivot em arqueivo excel 
pivot_table.to_excel("data/pivot_table.xlsx", "relatorio")