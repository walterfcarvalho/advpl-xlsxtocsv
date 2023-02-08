#include "rwmake.ch"
#include "protheus.ch"
#INCLUDE "font.ch"
#INCLUDE "colors.ch"
//#INCLUDE "inkey.ch"
#INCLUDE "topconn.ch"

//Marcos Teixeira. 14.03.2007
// 16.12.2009 - Adicionadas melhorias. Marcos Teixeira
// 17.12.2009 - Executa multi-linhas (separaçao por ; ) e trata ou não erros. Marcos Teixeira
// 22.12.2009 - Trata o retorno do tipo U, exibindo "-". Marcos Teixeira
// 06.01.2010 - Incluido verificação de fontes .PRX. Eduardo Carvalho
// 14.01.2010 - Correção comando altrim para alltrim - Ricardo Rauber
User Function AdvplExec()
    Local nWidth := 380
	local cFunc
	local bFlag	:= .t.
	local bRet	:= ""
	Private aFunc:= {'','','','','','','','','','','','','','','','','','','',''}

	If Select("SX2") == 0    // Conecto ao ambiente caso a execução seja feita a partir de um JOB.
		RPCSetType(3)
		RPCSetEnv("01","01",Nil,Nil,"LOJA")
	Endif

	DEFINE MSDIALOG oDlg FROM 0,0 TO 400,800 PIXEL TITLE "A D V P L    E X E C "
	DEFINE FONT oFnt NAME "Courier" Size 6,22

	oFont := TFont():New('Courier',,-12,.T.)
	oSay:= tSay():New(10,10,{||"Função | Fonte.PRW | INFO"},oDlg,,oFnt,,,,.T.)
	
    oMemo:= tMultiget():New(20,10,{|u|if(Pcount()>0,aFunc[pcount()+1]:=u,aFunc[1])},oDlg,nWidth,70,oFont,,,,,.T.)
	oMemo:appendText(MemoRead(GetClientDir() + "\advplexec.txt"))

	oSay:= tSay():New(095,10,{||"Retorno"},oDlg,,oFnt,,,,.T.)
	@ 105,10 MSGET oSay VAR bRet SIZE nWidth,50 OF oDlg PIXEL

	@ 170,010 BUTTON oButton PROMPT "OK(S/Err)"  OF oDlg PIXEL SIZE 40,15 ACTION {|| bRet:= "",  bRet:=Exc("") 	 , bFlag:=.t., oDlg:Refresh() }
	@ 170,050 BUTTON oButton PROMPT "OK(C/Err)"  OF oDlg PIXEL SIZE 40,15 ACTION {|| bRet:= "",  bRet:=ExcErr("") , bFlag:=.t., oDlg:Refresh() }
	@ 170,130 BUTTON oButton PROMPT "SAIR"  OF oDlg PIXEL ACTION {||oDlg:End() , bFlag:=.f.     }

	ACTIVATE MSDIALOG oDlg CENTERED

Return

//---------------------------------------------------
// Executa e trata algum erro, sfc.
Static Function Exc(xPar01)
	local n			:= 0
	local cR1
	local cResult
	Local cInfoComp
	Local cRes		:= ""
	Local aData		:= {}
	Local aInfoComp	:= {}
	Local oError := ErrorBlock({|e| MsgAlert("Ocorreu um Erro: " +chr(10)+ e:Description)})
	Local aFunc2	:= aClone(aFunc)

	cGrv:= ""
	for i:=1 to len(aFunc2)
		cGrv+= aFunc2[i]
	Next
	memowrite(GetClientDir() + "\advplexec.txt", cGrv)

	//Separa as linhas que tiverem ";"
	for n=1 to len(aFunc2)
		xPar01+=alltrim(aFunc2[n])
		aFunc2[n]:=""
	next
	if len(xPar01)>1
		xPar01+=";"
	endif
	//Remove as quebras de linha
	xPar01 := strtran(xPar01,chr(13),"")
	xPar01 := strtran(xPar01,chr(10),"")
	n:=1

	if ";" $ xPar01
		while (at(";",xPar01)>1)
			nPos:=at(";",xPar01)
			aFunc2[n]:=substr(xPar01,1,nPos-1)
			xPar01:=substr(xPar01,nPos+1,len(xPar01))
			n++
		enddo
	endif

	for n=1 to len(aFunc2)
		Begin Sequence
			if !empty(aFunc2[n])
				if (".PRW" $ upper(aFunc2[n])) .or. (".PRX" $ upper(aFunc2[n])) .and. (len(alltrim(aFunc2[n])) < 25)
					aData := GetAPOInfo(aFunc2[n])
					cMsg := "Nome do fonte: "+aData[1]
					cMsg += chr(13)+"Linguagem do fonte: "+aData[2]
					cMsg += chr(13)+"Modo de Compilação: "+aData[3]
					cMsg += chr(13)+"Ultima compilação do arquivo: "+dtoc(aData[4])
					cMsg += chr(13)+"Hora da compilação no RPO: "+aData[5]
					MsgInfo(cMsg)
				elseif upper(alltrim(aFunc2[n])) = "INFO"
					aInfoComp := GetRmtInfo()
					cInfoComp:= 'Nome do Computador: ' + aInfoComp[1]+chr(13)+chr(10)+;
					'Sistema Operacional: ' + aInfoComp[2]+chr(13)+chr(10)+;
					'Informação adicional: ' + aInfoComp[3]+chr(13)+chr(10)+;
					'Memória: ' + aInfoComp[4]+chr(13)+chr(10)+;
					'Nr. de Processadores: ' + aInfoComp[5]+chr(13)+chr(10)+;
					'MHZ Processador: ' + aInfoComp[6]+chr(13)+chr(10)+;
					'Descrição Processador: ' + aInfoComp[7]+chr(13)+chr(10)+;
					'Linguagem: ' + aInfoComp[8]+chr(13)+chr(10)+;
					'IP: '+GetClientIP()+chr(13)+chr(10)+;
					'Build: '+GetBuild()+chr(13)+chr(10)+;
					'Environment: '+GetEnvServer()+chr(13)+chr(10)+;
					'Tema: '+PtGetTheme()
					MsgInfo(cInfoComp)
				else
					cR1:=&("{|| "+alltrim(aFunc2[n])+" }")
					cResult:=eval(cR1)
					if valtype(cResult)="L"
						if cResult
							cRes+=".T."+chr(13)
						else
							cRes+=".F."+chr(13)
						endif
					elseif valtype(cResult)="N"
						cRes+=alltrim(str(cResult))+chr(13)
					elseif valtype(cResult)="D"
						cRes+=dToc(cResult)+chr(13)
					elseif valtype(cResult)="U"
						cRes+="-"+chr(13)
					else
						cRes+=cResult+chr(13)
					endif
				endif
			endif
		End Sequence
	Next

	Errorblock(oError)

Return(cRes)

//---------------------------------------------------
//Executa porém não trata erro
Static Function ExcErr(xPar01,xPar02)
	local n			:= 0
	local cR1
	local cResult
	Local cInfoComp
	Local cRes		:= ""
	Local aData		:= {}
	Local aInfoComp	:= {}
	Local aFunc2	:= aClone(aFunc)

	cGrv:= ""
	for i:=1 to len(aFunc2)
		cGrv+= aFunc2[i]
	Next
	memowrite(GetClientDir() + "\advplexec.txt",cGrv)

	//Separa as linhas que tiverem ";"
	for n=1 to len(aFunc2)
		xPar01+=alltrim(aFunc2[n])
		aFunc2[n]:=""
	next
	if len(xPar01)>1
		xPar01+=";"
	endif
	//Remove as quebras de linha
	xPar01 := strtran(xPar01,chr(13),"")
	xPar01 := strtran(xPar01,chr(10),"")
	n:=1

	if ";" $ xPar01
		while (at(";",xPar01)>1)
			nPos:=at(";",xPar01)
			aFunc2[n]:=substr(xPar01,1,nPos-1)
			xPar01:=substr(xPar01,nPos+1,len(xPar01))
			n++
		enddo
	endif

	for n=1 to len(aFunc2)
		Begin Sequence
			if !empty(aFunc2[n])
				if (".PRW" $ upper(aFunc2[n])) .or. (".PRX" $ upper(aFunc2[n])) .and. (len(alltrim(aFunc2[n])) < 25)
					aData := GetAPOInfo(aFunc[n])
					cMsg := "Nome do fonte: "+aData[1]
					cMsg += chr(13)+"Linguagem do fonte: "+aData[2]
					cMsg += chr(13)+"Modo de Compilação: "+aData[3]
					cMsg += chr(13)+"Ultima compilação do arquivo: "+dtoc(aData[4])
					cMsg += chr(13)+"Hora da compilação no RPO: "+aData[5]
					MsgInfo(cMsg)
				elseif upper(alltrim(aFunc2[n])) = "INFO"
					aInfoComp := GetRmtInfo()
					cInfoComp:= 'Nome do Computador: ' + aInfoComp[1]+chr(13)+chr(10)+;
					'Sistema Operacional: ' + aInfoComp[2]+chr(13)+chr(10)+;
					'Informação adicional: ' + aInfoComp[3]+chr(13)+chr(10)+;
					'Memória: ' + aInfoComp[4]+chr(13)+chr(10)+;
					'Nr. de Processadores: ' + aInfoComp[5]+chr(13)+chr(10)+;
					'MHZ Processador: ' + aInfoComp[6]+chr(13)+chr(10)+;
					'Descrição Processador: ' + aInfoComp[7]+chr(13)+chr(10)+;
					'Linguagem: ' + aInfoComp[8]+chr(13)+chr(10)+;
					'IP: '+GetClientIP()+chr(13)+chr(10)+;
					'Build: '+GetBuild()+chr(13)+chr(10)+;
					'Environment: '+GetEnvServer()+chr(13)+chr(10)+;
					'Tema: '+PtGetTheme()
					MsgInfo(cInfoComp)
				else
					cR1:=&("{|| "+alltrim(aFunc2[n])+" }")
					cResult:=eval(cR1)
					if valtype(cResult)="L"
						if cResult
							cRes+=".T."+chr(13)
						else
							cRes+=".F."+chr(13)
						endif
					elseif valtype(cResult)="N"
						cRes+=alltrim(str(cResult))+chr(13)
					elseif valtype(cResult)="D"
						cRes+=dToc(cResult)+chr(13)
					elseif valtype(cResult)="U"
						cRes+="-"+chr(13)
					elseif valtype(cResult)="A"
						cRes += "array"
					else
						cRes+=cResult+chr(13)
					endif
				endif
			endif
		End Sequence
	Next

Return(cRes)
