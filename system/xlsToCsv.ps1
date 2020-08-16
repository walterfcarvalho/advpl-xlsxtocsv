    #pega o diretorio
    $dir = ( Get-Item $args[0] ).DirectoryName
    
    #pega o nome base do arquivo
    $fileBase = (Get-Item $args[0] ).BaseName

    #define o nome do arquivo de saida
    $outFile = "$dir\$fileBase.csv"

    #operacoes com o excel
    $Excel = New-Object -ComObject Excel.Application

	$Excel.DisplayAlerts = $false
	$Excel.ScreenUpdating = $false
	$Excel.Visible = $false
	$Excel.UserControl = $false
	$Excel.Interactive = $false

    $excelFile = $args[0]
    
    $wb = $Excel.Workbooks.Open($excelFile)

	$workSheet = $wb.Sheets.Item(1);

    #deletar o arquivo se existir
    if( Test-Path $outFile ){
        Remove-Item $outFile     
    } 


	$workSheet.SaveAs($outFile, 06)

    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    spps -n Excel