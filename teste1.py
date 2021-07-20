import win32com.client as win32
import numpy as np
import pandas as pd

def openWorkbook(xlapp, xlfile):
    try:        
        xlwb = xlapp.Workbooks(xlfile)            
    except Exception as e:
        try:
            xlwb = xlapp.Workbooks.Open(xlfile)
        except Exception as e:
            print(e)
            xlwb = None                    
    return(xlwb)

try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = openWorkbook(excel, '\\Book1.xls') #aqui pode colocar a referencia do folder ou chamar por .bat em variavel
    ws = wb.Worksheets('Sheet1') 
    excel.Visible = False
    ws.PivotTables(1).PivotFields("UF").ClearAllFilters()
    ws.PivotTables(1).PivotFields("Produto").ClearAllFilters()

    uf_item = []
    produto_item = []
    item_to_remove = ['Sum of Vol', 'None', 'Grand Total', 'Ano', 'Mes']
    table_filter = []
    
    for item in ws.PivotTables(1).PivotFields("UF").PivotItems():
        uf = str(item)
        uf_item.append(uf)
        ws.PivotTables(1).PivotFields("UF").CurrentPage = uf
        for item in ws.PivotTables(1).PivotFields("Produto").PivotItems():
            produto = str(item)
            produto_item.append(produto)
            ws.PivotTables(1).PivotFields("Produto").CurrentPage = produto
            table_data = []
            counter = 0
            for item in ws.PivotTables(1).TableRange1:
                trange = str(item)
                counter = counter + 1
                if trange not in item_to_remove and counter == 5 or counter == 6 or counter == 7 or counter == 8 or counter == 9:

                    table_data.append(str(trange))

                    table_data.append(str(uf))
                    table_data.append(str(produto))

                    if counter == 9:
                        table_filter.append(table_data)
                        df = pd.concat([pd.Series(x) for x in table_filter], axis=1)

                        
    df1_transposed = df.T
    print(df1_transposed)
    #aqui pode automatizar o limite do indice das colunas para nao criar toda vez nominalmente cada frame
    data = [df1_transposed[1], df1_transposed[2], df1_transposed[0], df1_transposed[6], df1_transposed[9]]
            
    data2 = [df1_transposed[1], df1_transposed[2], df1_transposed[3], df1_transposed[6], df1_transposed[12]]
    
    df3 = pd.concat(data, axis=1)
    df4 = pd.concat(data2, axis=1)


    df3.columns = ['0', '1', '2', '3', '4']
    print(df3)
    df4.columns = ['0', '1', '2', '3', '4']
    print(df4)

    dfinal = pd.DataFrame(np.concatenate([df3.values, df4.values]), columns=df3.columns)
    print(dfinal) #a partir daqui seria formatação de datatypes e save .csv por ex

    wb.Save()


except Exception as e:
    print(e)

finally:
    # RELEASES RESOURCES
    ws = None
    wb = None
    excel = None

