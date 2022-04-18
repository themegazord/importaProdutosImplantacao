# Aplicação básica para importação de produtos e saldos com Python

## Status do projeto: Em desenvolvimento ⚠️

## Motivo da criação do projeto

### Na atual empresa onde trabalho, existe um sistema em Delphi com mais de 15 anos de vida, porém, nunca foi implementado um sistema importação de produtos e seus devidos saldos para a fase de implantação do sistema nos clientes, visto que, é a implantação que faz tal papel e não o desenvolvedor. Como no momento, estou atuando com Deployiment Analyst (Analista de Implantação) e não gosto muito de refazer etapas que eu vejo que um programa facilmente faria, resolvi criar um algoritmo em Python, juntamente com o Pandas para importação, via planilha Excel, produtos e seus saldos.

## Requisitos Minimos (os mesmo estão no arquivo ```requirements.txt```)

<table>
  <tr>
    <td>et-xmlfile</td>
    <td>firebirdsql</td>
    <td>numpy</td>
    <td>openpyxl</td>
    <td>pandas</td>
    <td>python-dateutil</td>
    <td>pytz</td>
    <td>six</td>
  </tr>
  <tr>
    <td>1.1.0</td>
    <td>1.2.1</td>
    <td>1.22.3</td>
    <td>3.0.9</td>
    <td>1.4.2</td>
    <td>2.8.2</td>
    <td>2022.1</td>
    <td>1.16.0</td>
  </tr>
</table> 

## Configurando o ```conexaobd.py```

### No meu caso, se tratou de um banco de dados especifico, sendo ele o Firebird.
### Para configurar o mesmo, você vai ter que preencher os seguintes parametros de ```firebirdsql.connect()```
### O seguintes parâmetros são:

1. user -> Caso não tenha trocado no seu firebird, o user vai permanecer o padrão, SYSDBA;
2. password -> Caso não tenha trocado no seu firebird, o password vai permanecer o padrão, masterkey;
3. database -> Aqui você irá inserir o caminho do arquivo do banco de dados, contendo lógico, a extensão ```.fdb```;
4. host -> Caso o servidor de execução esteja na sua máquina, use localhost, caso contrário, o nome ou o ip do servidor;
5. charset -> A configuração de texto do seu banco, por padrão, insira ANSI.


## Configurando o caminho do arquivo ```.xlsx```

### Para configurar o arquivo Excel, tanto na variavel ```leitor``` quanto na variavel ```leitorSaldo```, insira o caminho da mesma, visto que, caso a pasta esteja no mesmo local do programa, bastará apenas o nome do arquivo.

## Preenchendo os dados do arquivo Excel.

### Os dados dentro da planilha sempre, sem excessão, devem começar a partir da segunda linha, respeitando o nome das colunas.
### Cada campo segue as seguintes informações dentro do sistema:

### Todos os campos são obrigatório:

#### Sheet Estoque

1. NOME -> Nome do produto
2. REDUZIDO -> Nome reduzido do produto de até 10 caracteres.
3. COD_GRUPO -> Código do grupo a qual o produto pertence
4. CODIGOPRODUTO -> Código interno do produto
5. CODIGO_NCM -> Código NCM do produto
6. COD_SUBRUPO -> Código do subgrupo a qual o produto pertence
7. COD_CLASSE -> Código da classe a qual o produto pertence
8. COD_FABRICANTE -> Código do fabricante a qual o produto pertence
9. CEST -> Código CEST do produto, caso o mesmo esteja sendo tributado por substituição

#### Sheet Saldo

1. COD_PRODUTO -> Código do produto ao qual o saldo será inserido
2. COD_LOCAL -> Código do local de armazenamento do produto
3. PRECOVENDA -> Preço de venda do produto

## Execução do programa

### Com todos os arquivos configurados, basta executar o ```captaprodutos.py``` e um log dos produtos irá aparecer no CMD se foram ou não inseridos no banco de dados.
