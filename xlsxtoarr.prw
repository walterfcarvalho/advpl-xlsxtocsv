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


User function xlsxToArr(cArq, cIdPlan, cDelimiter, cSaltoLinhas)
    Local oProcess  := nil
    Local aRes      := nil
    Local lEnd      := .F.

    Default cIdPlan := "1"
    Default cArq    := ""
    Default cDelimiter := ","
    Default cSaltoLinhas := "0"

    oProcess := MsNewProcess():New({|lEnd| aRes:= Converter(cArq, cIdPlan, cDelimiter, cSaltoLinhas, @oProcess, @lEnd)  },"Extraindo dados da planilha XLSX","Efetuando a leitura do arquivo xlsx...", .T.)

    oProcess:Activate()

Return aRes

Static Function Converter(cArq, cIdPlan, cDelimiter, cSaltoLinhas, oProcess, lEnd)
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
    Local aLinha    := {}
    Local lManterVazio := .T.  

    //setar o delimitador  
    If cDelimiter <> ";"
       cDelimiter := ","		
    EndIf

    //Testar se existe excel instalado na maquina
     If ApOleClient("MsExcel") = .F.
        ApMsgStop("Nao detectei excel instalado na maquina:", cMsgHead)
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
    cArqCsv := GetTempPath() + zArqSemExt(cArq) + ".csv"
    cArqTmp := GetTempPath() + zArqSemExt(cArq) + ".tmp"

    FErase(cArqCsv)
    FErase(cArqTmp)

    // Valida se o arquivo informado existe
    If File(cArq,/*nWhere*/,.T.) = .F.
        ApMsgStop("Arquivo n�o encontrado:" + cArq, cMsgHead)
        Return aRes
    EndIf

    oProcess:SetRegua1(4)
    oProcess:SetRegua2(2)

    oProcess:IncRegua1("1/4 Baixar xlsxTocsv.exe")
    oProcess:IncRegua2("")

    // Pega do servidor o arquivo que vai converter o xlsx  para csv
    If CpyS2T("\system\xlsxtocsv.exe", GetClientDir(), .F., .F.) = .F.
        ApMsgStop('Nao foi possivel baixar o conversor do servidor, em "\system\"' + cExe, cMsgHead)
        Return aRes
    EndIf

    oProcess:IncRegua1("2/4 Arq CSV temporario")
    oProcess:SetRegua2(20)

    nShell := Shellexecute('open', '"' + GetClientDir() + cExe + '"', '"' + Alltrim(cArq) + '" "' + cIdPlan + '" "' + cDelimiter + '" "' + cSaltoLinhas + '" ' , GetClientDir(), 0)

    While File(cArqCsv) = .F.
        nPassos += 1

        if lEnd = .T.    //VERIFICAR SE N�O CLICOU NO BOTAO CANCELAR
            ApMsgStop("Processo cancelado pelo usuario." + cArq, cMsgHead)
            Return aRes
        EndIf

        If nPassos = 50
            ApMsgStop("A conversso excedeu o tempo limite para o arquivo" + cArq, cMsgHead)
            Return aRes
        EndIf

        oProcess:IncRegua2("Convertendo arquivo...")

        If nShell = -1 .Or. nShell = 2
            ApMsgStop("Nao foi possivel efetuar a conversao do arquivo." + cArq, cMsgHead)
            Return aRes
        Else
            Sleep(1000)
        EndIf

    EndDo

    oFile := FWFileReader():New(cArqCsv)

    If oFile:Open() = .F.
        ApMsgStop("Nao foi possivel efetuar a leitura do arquivo." + cArq, cMsgHead)
        Return aRes
    EndIf

    aLines := oFile:GetAllLines()

    if lEnd = .T.   //VERIFICAR SE N�O CLICOU NO BOTAO CANCELAR
        ApMsgStop("Processo cancelado pelo usuario." + cArq, cMsgHead)
        Return aRes
    EndIf


    oProcess:IncRegua1("3/4 Ler Arquivo CSV")
    oProcess:SetRegua2(Len(aLines))

    For i:=1 to len(aLines)

        if lEnd = .T.    //VERIFICAR SE N�O CLICOU NO BOTAO CANCELAR
            ApMsgStop("Processo cancelado pelo usuario." + cArq, cMsgHead)
            Return {}
        EndIf

        oProcess:IncRegua2("Lendo registro " + CvalToChar(i) + " de " + cValToCHar(Len(aLines)) )

        cLinha  := aLines[i]

        If Empty(cLinha) = .F.
            cLinha := StrTran(cLinha, '"', '')

            aLinha := Separa(cLinha, ",", lManterVazio)

            If Len(aLinha) > 0
                Aadd( aRes, aLinha )
            EndIf    
        EndIf
    Next

    oFile:Close()

    oProcess:IncRegua1("4/4 Remove temporarios")
    oProcess:SetRegua2(1)
    oProcess:IncRegua2("")

    FErase(cArqCsv)
    FErase(cArqTmp)

Return aRes

 Static Function zArqSemExt(cArq)
    Local i          := 1
    Local nPosPonto  := 1
    Local nPosBarra  := 1

    For i:= Len(cArq) to 1 step -1
        if Substr(cArq, i) = "."
            nPosPonto := i
            Exit
        EndIf
    Next

    For i:= Len(cArq) to 1 step -1
        if Substr(cArq, i) = "\"
            nPosBarra := i
            Exit
        EndIf
    Next    

Return  Substr(cArq, nPosBarra + 1,    nPosPonto - nPosBarra -1  )
