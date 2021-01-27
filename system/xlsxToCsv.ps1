    #pegar o idx da planilha se for passado o parametro, senão pega a primeira
    if ( $args[1] ){
        $nPlanilha = [int]$args[1]
    } Else {
        $nPlanilha = 1
    }

    #pega o diretorio
    #$dir = ( Get-Item $args[0] ).DirectoryName
    
    #pega o nome base do arquivo
    $fileBase = (Get-Item $args[0] ).BaseName

    #define o nome do arquivo de saida
    $outFile = "$env:TEMP\$fileBase.tmp"

    #deletar o arquivo tmp, se existir
    if( Test-Path $outFile ){
        Remove-Item $outFile     
    } 
    
    #deletar o arquivo csv, se existir
    if( Test-Path "$env:TEMP\$fileBase.csv" ){
        Remove-Item "$env:TEMP\$fileBase.csv"    
    } 


    #operacoes com o excel
    $Excel = New-Object -ComObject Excel.Application

	$Excel.DisplayAlerts = $false
	$Excel.ScreenUpdating = $false
	$Excel.Visible = $false
	$Excel.UserControl = $false
	$Excel.Interactive = $false

    $excelFile = $args[0]
    
    $wb = $Excel.Workbooks.Open($excelFile)

	$workSheet = $wb.Sheets.Item($nPlanilha);
    
    #salva no formato csv
	$workSheet.SaveAs($outFile, 6)

    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    spps -n Excel

    ECHO "$env:TEMP\$fileBase.csv"
    ECHO $outFile

    #reprocessa com o delimitador escolhido
    Import-Csv $outFile | Export-Csv "$env:TEMP\$fileBase.csv" -Delimiter $args[2] -NoTypeInformation

    #deletar o arquivo tmp, se existir
    if( Test-Path $outFile ){
        Remove-Item $outFile     
    } 
    