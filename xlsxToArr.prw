//-------------------------------------------------------------------
/*/{Protheus.doc}

@description    Efetua a conversao para csv da primeira planilha passada como parametro
@author         @walterfcarvalho
@since          16/08/2020
@version        1.00
/*/
//-------------------------------------------------------------------


User function xlsxToArr(cArq)
    Local aRes := nil

    FWMsgRun(, {|| aRes :=  Converter(cArq) }, "", "Efetuando a leitura do arquivo xlsx... ")

Return aRes

Static  function Converter(cArq)
    Local aRes    := {}
    Local nHandle := 0
    Local cExe    := "xlsxToCsv.exe"
    Local cArqCsv := SubStr(cArq, 1, Len(cArq) -3 ) + "csv"

    // Se nao enviar cArq, abre dialogo para escolher o arquivo
    If Empty(cArq) = .T.
        cArq := cGetFile( "*.xls", "Informe o Arquivo XLS", 1, "c:\", .F., GETF_LOCALHARD, .T., .T. )
    EndIf

    // Valida se o arquivo informado existe
    If File(cArquivo,/*nWhere*/,.T.) = .F.
        MsgBox("Arquivo não encontrado:" + cArq, cCabeca , "STOP")
        Return aRes
    EndIf

    // Pega do servidor o arquivo que vai converter o xlsx  para csv
    If File( GetClientDir() + cExe ) = .F.
        If CpyS2T("\system\xlsxToCsv.exe", GetClientDir(), .F., .F.)
            MsgBox('Não foi possível pegar o conversor no servidor, em "\system\"' + cExe, xlsxToArr(), "STOP")
            Return aRes
        EndIf
    EndIf

    // Chamar o executavel externo para gerar um CSV
    WaitRun( GetClientDir() + cExe + " " + cArq, SW_SHOWNORMAL )

    If File(cArqCsv) = .F.
        MsgBox("Não foi possivel efetuar a conversão do arquivo." + cArq, xlsxToArr(), "STOP")
        Return aRes
    EndIf

    nHandle := FT_FUse(cArqCsv)
    If nHandle < 0
        MsgBox("Não foi possivel ler o arquivo CSV." + cArq, cCabeca, "STOP")
        Return aRes
    EndIf

    // Posiciona na primeria linha
    FT_FGoTop()

    While !FT_FEOF()
        cLinha  := FT_FReadLn()

        Aadd( aRes, Separa(cLinha, ",", .F.))

        FT_FSKIP()
    EndDo

    // Fecha o Arquivo
    FT_FUSE()

    // remove o arquivo csv
    FErase(cArqCsv)

Return aRes