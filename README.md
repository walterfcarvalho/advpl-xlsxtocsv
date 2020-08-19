# advpl-xlsxtocsv
Projeto com a classe conversora de xlsx para csv

# Como funciona?

O executável xlsxToCsv.exe é um script de powershell que é encapsulado dentro de um executável escrito em C#, como visto em 

https://purainfo.com.br/convertendo-powershell-scripts-em-executaveis-ps1-para-exe/


Então o fonte xlsxToArr.prw faz uma chamdada para xlsxToCsv.exe via Shellexecute(), que gera um arquivo CSV temporário, que então pode ser lido nativamente pelo protheus.

# Como Fazer o deploy?

Basta colocar o executável xlsxToCsv.exe na pasta system

E compilar o fonte xlsxToArr.prw  no seu projeto ;)
