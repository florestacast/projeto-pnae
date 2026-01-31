import openpyxl
import psycopg2
from psycopg2.extras import execute_values
import re
import os
from dotenv import load_dotenv

load_dotenv()

DB_CONFIG = {
    'host': os.getenv('DB_HOST', '127.0.0.1'),
    'port': int(os.getenv('DB_PORT', 5433)),
    'database': os.getenv('DB_NAME', 'florestacast'),
    'user': os.getenv('DB_USER', 'florestacast'),
    'password': os.getenv('DB_PASS', '')
}

ARQUIVO_EXCEL = r'D:\Meu Drive\Trabalho\Floresta_cast\Clientes\Luzinia\Envelope\Projeto\projeto_venda.xlsx'
PROJETO_ID = 1

def conectar_banco():
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        return conn
    except Exception as e:
        print(f'Erro: {e}')
        exit(1)

def limpar_valor(valor, tipo):
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return valor
    valor_str = str(valor).strip()
    if tipo == 'quantidade':
        match = re.search(r'(\d+)', valor_str)
        return int(match.group(1)) if match else 0
    if tipo == 'preco' or tipo == 'total':
        valor_limpo = re.sub(r'[^\d,.]', '', valor_str)
        valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
        try:
            return float(valor_limpo)
        except:
            return 0.0
    return valor_str

def importar_dados_excel():
    try:
        workbook = openpyxl.load_workbook(ARQUIVO_EXCEL)
        planilha = workbook.active
        print(f'Lendo: {planilha.title}')
        print('=' * 80)
        conn = conectar_banco()
        cur = conn.cursor()
        dados = []
        for linha_idx in range(2, planilha.max_row + 1):
            itens = limpar_valor(planilha[f'A{linha_idx}'].value, 'quantidade')
            produto = str(planilha[f'B{linha_idx}'].value or '').strip()
            unidade = str(planilha[f'C{linha_idx}'].value or '').strip()
            quantidade = limpar_valor(planilha[f'D{linha_idx}'].value, 'quantidade')
            preco = limpar_valor(planilha[f'E{linha_idx}'].value, 'preco')
            total = limpar_valor(planilha[f'F{linha_idx}'].value, 'total')
            sazonalidade = str(planilha[f'G{linha_idx}'].value or '').strip()
            status_alimentos = str(planilha[f'H{linha_idx}'].value or '').strip()
            if produto:
                dados.append((PROJETO_ID, itens, produto, unidade, quantidade, preco, total, sazonalidade, status_alimentos))
        if dados:
            sql = 'INSERT INTO alimentos (projeto_id, itens, produto, unidade, quantidade, preco, total, sazonalidade, status_alimentos) VALUES %s'
            execute_values(cur, sql, dados)
            conn.commit()
            print(f'✓ {len(dados)} linhas importadas!')
        print('\nRELATÓRIO:')
        cur.execute('SELECT COUNT(*) FROM alimentos WHERE projeto_id = %s', (PROJETO_ID,))
        print(f'Total: {cur.fetchone()[0]} itens')
        cur.execute('SELECT SUM(total) FROM alimentos WHERE projeto_id = %s', (PROJETO_ID,))
        valor = cur.fetchone()[0] or 0
        print(f'Valor: R\$ {valor:,.2f}')
        cur.close()
        conn.close()
        workbook.close()
    except Exception as e:
        print(f'Erro: {e}')
        exit(1)

if __name__ == '__main__':
    importar_dados_excel()
    print('Concluído!')
