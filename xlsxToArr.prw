//-------------------------------------------------------------------
/*/{Protheus.doc}

@description    Efetua a conversao para csv da primeira planilha passada como parametro
@author         @walterfcarvalho
@since          16/08/2020
@version        1.00
/*/
//-------------------------------------------------------------------

#include "shell.ch"


User function xlsxToArr(cArq)
    Local oProcess  := nil
    Local aRes      := nil
    Local lEnd      := .F.
    
    lEnd      := .F.

	oProcess := MsNewProcess():New({|lEnd| aRes:= Converter(cArq, @oProcess)  },"Extraindo dados da planilha XLSX","Efetuando a leitura do arquivo xlsx...", .F.)

	oProcess:Activate()

Return aRes

Static Function Converter(cArq, oProcess)
    Local nReg      := 0    
    Local nShell    := 0
    Local cMsgHead  := "xlsxToArr()"
    Local cDirIni   := StrTran(GetTempPath(), "AppData\Local\Temp\", "")
    Local aRes      := {}
    Local nHandle   := 0
    Local cExe      := "xlsxToCsv.exe"
    Local cArqCsv   := ""
    Local cArqTmp   := ""


    // Se nao enviar cArq, abre dialogo para escolher o arquivo
    If Empty(cArq) = .T.
        cArq := cGetFile( "Arquivos Excel|*.xlsx", "Informe o Arquivo XLSX",  1, cDirIni, .F., GETF_LOCALHARD, .F., .T. )
    EndIf

    // Gere o nome do arquivo CSV temporario
    cArqCsv := SubStr(cArq, 1, Len(cArq) -4 ) + "csv"
    cArqTmp := SubStr(cArq, 1, Len(cArq) -4 ) + "tmp"

    // Valida se o arquivo informado existe
    If File(cArq,/*nWhere*/,.T.) = .F.
        ApMsgSTop("Arquivo não encontrado:" + cArq, cMsgHead)
        Return aRes
    EndIf

    oProcess:SetRegua1(4)
    oProcess:SetRegua2(2)
    
    oProcess:IncRegua1("1/4 Baixar xlsxTocsv.exe")
    oProcess:IncRegua2("")

    // Pega do servidor o arquivo que vai converter o xlsx  para csv
    If File( GetClientDir() + cExe ) = .F.
        If CpyS2T("\system\xlsxtocsv.exe", GetClientDir(), .F., .F.) = .F.
            ApMsgSTop('Não foi possível pegar o conversor no servidor, em "\system\"' + cExe, cMsgHead)
            Return aRes
        EndIf

    EndIf

    oProcess:IncRegua1("2/4 Arq CSV temporario")
    oProcess:SetRegua2(10)

    nShell := Shellexecute('open', '"' + GetClientDir() + cExe + '"', '"' + Alltrim(cArq) + '"', GetClientDir(), 2)

    While File(cArqCsv) = .F.
        oProcess:IncRegua2("Convertendo arquivo...")

        If nShell = -1 .Or. nShell = 2
            ApMsgSTop("Não foi possivel efetuar a conversão do arquivo." + cArq, xlsxToArr())
            Return aRes
        Else    
            Sleep(1000)
        EndIf 

    EndDo

    nHandle := FT_FUse(cArqCsv)
    If nHandle < 0
        ApMsgSTop("Não foi possível ler o arquivo CSV." + cArq, cMsgHead)
        Return aRes
    EndIf


    oProcess:IncRegua1("3/4 Ler Arquivo CSV")
    oProcess:SetRegua2(FT_FLastRec())

    // Posiciona na primeria linha
    FT_FGoTop()

    While !FT_FEOF()
        nReg += 1
        oProcess:IncRegua2("Lendo registro " + CvalToChar(nReg) + " de " + cValToCHar(FT_FLastRec()) )

        cLinha  := FT_FReadLn()

        If Empty(cLinha) = .F.    
            Aadd( aRes, Separa(cLinha, ",", .F.))
        EndIf            

        FT_FSKIP()
    EndDo


    oProcess:IncRegua1("4/4 Remove temporarios")
    oProcess:SetRegua2(1)
    oProcess:IncRegua2("")

    // Fecha o Arquivo
    FT_FUSE()

    // remove o arquivo csv
    FErase(cArqCsv)
    FErase(cArqTmp)

Return aRes
