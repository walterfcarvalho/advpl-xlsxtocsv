//-------------------------------------------------------------------
/*/{Protheus.doc}

@description    Efetua a conversao para csv da primeira planilha passada como parametro
@author         @walterfcarvalho
@since          16/08/2020
@version        1.00
/*/
//-------------------------------------------------------------------

#INCLUDE 'PROTHEUS.CH'
#include "shell.ch"
#INCLUDE 'TOTVS.CH'


User function xlsxToArr(cArq, cIdPlan)
    Local oProcess  := nil
    Local aRes      := nil
    Local lEnd      := .F.

    Default cIdPlan := "1"
    Default cArq    := ""

    oProcess := MsNewProcess():New({|lEnd| aRes:= Converter(cArq, cIdPlan, @oProcess, @lEnd)  },"Extraindo dados da planilha XLSX","Efetuando a leitura do arquivo xlsx...", .T.)

    oProcess:Activate()

Return aRes

Static Function Converter(cArq, cIdPlan, oProcess, lEnd)
    Local i         := 1
    Local aLines    := {}
    Local oFile     := Nil
    Local nPassos   := 0
    Local nShell    := 0
    Local cMsgHead  := "xlsxToArr()"
    Local cDirIni   := StrTran(GetTempPath(), "AppData\Local\Temp\", "")
    Local aRes      := {}
    Local cExe      := "xlsxToCsv.exe"
    Local cArqCsv   := ""
    Local cArqTmp   := ""

    //Testar se existe excel instalado na maquina
     If ApOleClient("MsExcel") = .F.
        ApMsgStop("Não detectei excel instalado na máquina:", cMsgHead)
        Return aRes
    EndIf
    
    // Se nao enviar cArq, abre dialogo para escolher o arquivo
    If Empty(cArq) = .T.
        cArq := cGetFile( "Arquivos Excel|*.xlsx|Arquivos Excel 97|*.xls", "Selecione o arquivo:",  1, cDirIni, .F., GETF_LOCALHARD, .F., .T. )

        If Empty(cArq)
            ApMsgStop("Importacao Cancelada:", cMsgHead)
            Return aRes
        EndIf
    EndIf

    // Gere o nome do arquivo CSV temporario
    cArqCsv := GetTempPath() + RetFileName(cArq) + ".csv"
    cArqTmp := GetTempPath() + RetFileName(cArq) + ".tmp"

    // Valida se o arquivo informado existe
    If File(cArq,/*nWhere*/,.T.) = .F.
        ApMsgStop("Arquivo não encontrado:" + cArq, cMsgHead)
        Return aRes
    EndIf

    oProcess:SetRegua1(4)
    oProcess:SetRegua2(2)

    oProcess:IncRegua1("1/4 Baixar xlsxTocsv.exe")
    oProcess:IncRegua2("")

    // Pega do servidor o arquivo que vai converter o xlsx  para csv
    If File( GetClientDir() + cExe ) = .F.
        If CpyS2T("\system\xlsxtocsv.exe", GetClientDir(), .F., .F.) = .F.
            ApMsgStop('Não foi possível baixar o conversor do servidor, em "\system\"' + cExe, cMsgHead)
            Return aRes
        EndIf

    EndIf

    oProcess:IncRegua1("2/4 Arq CSV temporario")
    oProcess:SetRegua2(20)

    nShell := Shellexecute('open', '"' + GetClientDir() + cExe + '"', '"' + Alltrim(cArq) + '" "' + cIdPlan + '" ' , GetClientDir(), 0)

    While File(cArqCsv) = .F.
        nPassos += 1

        if lEnd = .T.    //VERIFICAR SE NÃO CLICOU NO BOTAO CANCELAR
            ApMsgStop("Processo cancelado pelo usuário." + cArq, cMsgHead)
            Return aRes
        EndIf

        If nPassos = 50
            ApMsgStop("A conversão excedeu o tempo limite para o arquivo" + cArq, cMsgHead)
            Return aRes
        EndIf

        oProcess:IncRegua2("Convertendo arquivo...")

        If nShell = -1 .Or. nShell = 2
            ApMsgStop("Não foi possível efetuar a conversão do arquivo." + cArq, cMsgHead)
            Return aRes
        Else
            Sleep(1000)
        EndIf

    EndDo

    oFile := FWFileReader():New(cArqCsv)

    If oFile:Open() = .F.
        ApMsgStop("Não foi possível efetuar a leitura do arquivo." + cArq, cMsgHead)
        Return aRes
    EndIf

    aLines := oFile:GetAllLines()

    if lEnd = .T.   //VERIFICAR SE NÃO CLICOU NO BOTAO CANCELAR
        ApMsgStop("Processo cancelado pelo usuário." + cArq, cMsgHead)
        Return aRes
    EndIf


    oProcess:IncRegua1("3/4 Ler Arquivo CSV")
    oProcess:SetRegua2(Len(aLines))

    For i:=1 to len(aLines)

        if lEnd = .T.    //VERIFICAR SE NÃO CLICOU NO BOTAO CANCELAR
            ApMsgStop("Processo cancelado pelo usuário." + cArq, cMsgHead)
            Return {}
        EndIf

        oProcess:IncRegua2("Lendo registro " + CvalToChar(i) + " de " + cValToCHar(Len(aLines)) )

        cLinha  := aLines[i]

        If Empty(cLinha) = .F.
            Aadd( aRes, Separa(cLinha, ",", .F.))
        EndIf
    Next

    oFile:Close()

    oProcess:IncRegua1("4/4 Remove temporarios")
    oProcess:SetRegua2(1)
    oProcess:IncRegua2("")

    FErase(cArqCsv)
    FErase(cArqTmp)

Return aRes
