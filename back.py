import sqlite3
banco = sqlite3.connect('Art_Doce.db')
cursor = banco.cursor()
def read_task():
    c = banco.cursor()
    c.execute('''SELECT* FROM Art_Doce''')
    data = c.fetchall()
    banco.commit()
    return data


''' nome = 'Jean'
        senha = '070199'
        nome1 = 'Jéssica'
        senha1 = '280916'
        if values ['usuario'] == nome and values ['senha'] == senha:
            janela1.hide()
            janela2 = janela_consulta()
        elif values['usuario'] == nome1 and values ['senha'] == senha1:
            janela1.hide()
            janela2=janela_consulta()'''