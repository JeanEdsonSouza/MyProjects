import openpyxl as op

def criar_tabela():
    while True:        
        a = str(input("Digite o nome da Planilha: "))
        bk = op.Workbook()
        bk.create_sheet(f"{a}")
        chepe_page = bk[f'{a}']
        cat = str(input("Digite o nome da Aba: "))
        chepe_page.append([f'{cat}'])
        ad = int(input("deseja adcionar mais 1 categoria: [1] para sim [2] para não"))
        if ad ==1:
            cat1 = str(input("Digite o nome da Aba: "))
            chepe_page.append([f'{cat}',f'{cat1}'])
            bk.save(f"{a}.xlsx")
            break

        else:
            bk.save(f"{a}.xlsx")
            break
criar_tabela()