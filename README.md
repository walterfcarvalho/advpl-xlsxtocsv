
# advpl-xlsxtocsv
Projeto com a classe conversora de xlsx para csv

# Como funciona?

O executável xlsxToCsv.exe é um script de powershell que é encapsulado dentro de um executável escrito em C#, como visto em 
https://purainfo.com.br/convertendo-powershell-scripts-em-executaveis-ps1-para-exe/


Então o fonte xlsxToArr.prw faz uma chamdada para xlsxToCsv.exe via Shellexecute(), que  consegue ler arquivos XLSX ou XLS, gerar um arquivo CSV temporário, que então pode ser lido nativamente pelo protheus. 

Após isso a funcção devolve o conteúdo da planilha em formato de array.

# Como Fazer o deploy?

Basta colocar o executável xlsxToCsv.exe na pasta system

E compilar o fonte xlsxToArr.prw  no seu projeto ;)

# Como funciona?
Após a deploy use u_xlsxToArr() com dois parâmetros

cArq: nome do arquivo XLSX ou XLS. Se não for passado, então a função abre uma janela para escolher o arquivo.

cId: id da planilha no arquivo. Se não for informado, será considerada a primeira planilha.  

cSeparator: Permite separar os campos do CSV com ',' ou ";". Se não for informado cria o arquivo separado por ','.

# importante 
Para serem importadas todas as colunas, é necessário que na linha 01, seja inserido algum valor para cada coluna que deseja importar,
(pois  para verificar dinamicamente quais celulas são preenchidas, demorariaa demais e ninguém quer isso, correto ?)


Minhas redes sociais: 
twitter: @walterfcarvalho
inta: /walterfcarvalho
linkedin: /walterfcarvalho

Se ainda sim tiver dificuldades, manda um wmail para mim em walterfcarvalho@gmail.com ;)
