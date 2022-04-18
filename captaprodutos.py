import pandas as pd

from conexaobd import conn

cur = conn.cursor()

temp = dict()
tempSaldo = dict()
leitor = pd.read_excel("Importação de Cadastro de Produtos.xlsx", sheet_name='Estoque')
leitorSaldo = pd.read_excel("Importação de Cadastro de Produtos.xlsx", sheet_name='Saldo')
qtdLinhs = leitor.shape[0]
qtdLinhasSaldo = leitorSaldo.shape[0]

for l in range(0, qtdLinhs):
    temp = {
        'NOME' : f'{leitor.loc[l, "NOME"]}',
        'REDUZIDO' : f'{leitor.loc[l, "REDUZIDO"]}',
        'COD_GRUPO' : f'{leitor.loc[l, "COD_GRUPO"]}',
        'CODIGOPRODUTO' : f'{leitor.loc[l, "CODIGOPRODUTO"]}',
        'CODIGO_NCM' : f'{leitor.loc[l, "CODIGO_NCM"]}',
        'COD_SUBGRUPO' : f'{leitor.loc[l, "COD_SUBGRUPO"]}',
        'COD_CLASSE' : f'{leitor.loc[l, "COD_CLASSE"]}',
        'COD_FABRICANTE' : f'{leitor.loc[l, "COD_FABRICANTE"]}',
        'CEST' : f'{leitor.loc[l, "CEST"]}',
    }
    #Insert do estoque
    cur.execute(f'INSERT INTO EST_PRODUTO (COD_PRODUTO, NOME, REDUZIDO, TRIBUTACAO, BLOQUEADO, CLASSIFISCAL, MENSAGEM, COD_GRUPO, COD_FAMILIA, COD_UNIDADE, COD_FORMULA, ORIGEM, CODIGOBARRAS, COD_UNIDADEEMB, CODIGOPRODUTO, CLASSIFICACAO_MA, CODIGO_NCM, PESOUNIDADE, LISTATABELAPRECO, COD_SUBGRUPO, COD_CLASSE, COD_CADASTROFOR, COD_FABRICANTE, DATAATUALIZACAO, TIPO_SPED, GENERO, CONTROLE_LOTE, COD_IMPOSTO, EXPORTA_BALANCA, COMPOSTO, QTDE_MULTIPLA, TRANSPASSE, DESCRICAO_DETALHADA, IMPRIMECOMPOSICAO, UTILIZAPRECOCOMPOSICAO, CODIGO_INTEGRACAO, SOMENTE_ORCAMENTO, REFERENCIA_BARRAS, CEST, TEMVALIDADE, COD_DECRETO_MENS, TEMGRADE, LISTACATALOGO, BASEDETROCA, BALANCA_VENDA_UND, CODIGO_ANP, DESCRICAO_ANP, ALTURA, LARGURA, PROFUNDIDADE, SKU, DESCRICAO_ECOM, PALAVRAS_ECOM, TITULO_ECOM, ECOMMERCE) VALUES(gen_id(cod_produtogen, 1), \'{temp["NOME"]}\', \'{temp["REDUZIDO"]}\', \'T\', \'N\', NULL, NULL, {temp["COD_GRUPO"]}, NULL, 1, NULL, \'0\', NULL, 1, {temp["CODIGOPRODUTO"]},NULL,{temp["CODIGO_NCM"]}, 0, \'S\', {temp["COD_SUBGRUPO"]}, {temp["COD_CLASSE"]}, NULL, {temp["COD_FABRICANTE"]}, NULL, \'00\', NULL, \'N\', NULL, \'N\', \'N\', NULL, NULL, NULL, \'N\', \'N\', NULL, \'N\', NULL, \'{temp["CEST"]}\', \'N\', NULL, \'N\', \'N\', \'N\', \'P\', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, \'N\');')
    print(f'Produto {temp["NOME"]} cadastrado')
    del(temp)


for s in range(0, qtdLinhasSaldo):
    tempSaldo = {
        'COD_PRODUTO' : f'{leitorSaldo.loc[s, "COD_PRODUTO"]}',
        'COD_LOCAL' : f'{leitorSaldo.loc[s, "COD_LOCAL"]}',
        'PRECOVENDA' : f'{leitorSaldo.loc[s, "PRECOVENDA"]}',
    }
    cur.execute(f'INSERT INTO EST_SALDO (COD_SALDO, COD_PRODUTO, COD_LOCAL,CUSTOMEDIO,PEDIDOIDEAL, CUSTOFIXO, CUSTOVARIAVEL, MARGEMLUCRO,  BLOQ_INVENTARIO, PRECOVENDA, BLOQUEADO_LOCAL) VALUES (gen_id(cod_saldogen, 1),{tempSaldo["COD_PRODUTO"]},{tempSaldo["COD_LOCAL"]},0,0,0,0,0,\'N\',{tempSaldo["PRECOVENDA"]}, \'N\');')
    print(f'O saldo do produto {tempSaldo["COD_PRODUTO"]} foi inserido')
    del(tempSaldo)

cur.execute('COMMIT WORK;')
print('Comitado')

