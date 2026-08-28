#Include "Protheus.ch"
#INCLUDE "TOPCONN.CH"

//pablo :  use o EST033.tlpp como modelo de query e colunas do excel = copiar e DbCloseArea

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ EST051    ºAutor  ³ Vinicius        º Data ³  26/02/2025   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Compras - Relação Ordem de Produção vencendo a 5 dias      º±±
±±º          ³------------------------------------------------------------º±±
±±º          ³ Lê o SC2-Ordens de Producao, mostra todos com status em    º±±
±±º          ³ aberto e pendentes de entrega para os próximos 5 dias.     º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºRegra     ³ Uso exclusivo Prolinhas, envia e-mailconforme regra:       º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function EST051(xRotAuto)   
	Local aMensagem 	:= {}
	Local aBotoes   	:= {}
	Private cTitulo   	:= "EST051 - Movimentos Internos Custeados"
	Private lRotAuto	:= If(xRotAuto==Nil,.F.,xRotAuto)
	Private aDados 		:= {}                     
   	Private lProLinhas := If(Substring(CNUMEMP,1,2) == "03",.T.,.F. )  //NAO PRECISO DISSO PABLO
	
	If lRotAuto      
	    ConOut( cTitulo+" - "+Alltrim(SM0->M0_CODIGO)+" "+Alltrim(SM0->M0_CODFIL)+" - "+Alltrim(SM0->M0_FILIAL) )
		RptDetail()
	Else	
		AAdd( aMensagem, "Este programa tem como objetivo gerar uma apresentação em MS-EXCEL")
		AAdd( aMensagem, "Lê o SD3-Movimentos Internos Estoque, mostra todos os lançamentos  ")
		AAdd( aMensagem, "filtrados, exceto as transferências.")
		AAdd( aMensagem, "")
	
		AAdd( aBotoes, { 01, .T., { || SX5Param(), FechaBatch() } } )
		AAdd( aBotoes, { 22, .T., { || FechaBatch() } } )
	
		//FormBatch(<cTitulo>,<aMensagem>,<aBotoes>,<bValid>,nAltura,nLargura)
		FormBatch( cTitulo, aMensagem, aBotoes, , 160 )
	Endif	
Return Nil
               
**************************
Static Function SX5Param() 
	If ApOleClient("MsExcel")
		Processa( {|| RptDetail()}, cTitulo, "Processando arquivo, aguarde...", .F. )                     
	Else
		MsgInfo("MsExcel não instalado")
	Endif
Return Nil

***************************
Static Function RptDetail()
	// variaveis do XML/Excel
   	Local oExcel 	:= FwMsExcelEx():New()
   	Local cArq 		:= ""
   	Local cDir 		:= GetSrvProfString("Startpath","") 
   	//Local cDir 		:= GetClientDir()
	Local cWorkSheet:= ""
	Local cTable 	:= ""
 	
	// variaveis da rotina
	Local nReg 		:= 0
	Local _dtVencto := Date()+5
	Local cQuery	:= ""
	Local x

	//-- gera TIPO "MS"  e depois TIPO not in "MS,MP"
	For x:= 1 to 1  //pablo retirar esse conceito
		cQuery:= "SELECT" + CRLF
		cQuery+= "    C2_FILIAL," + CRLF
		cQuery+= "    C2_NUM + C2_ITEM + C2_SEQUEN AS C2NUM," + CRLF
		cQuery+= "    C2_EMISSAO," + CRLF
		cQuery+= "    C2_DATPRI," + CRLF
		cQuery+= "    C2_DATPRF," + CRLF
		cQuery+= "    C2_PRODUTO," + CRLF
		cQuery+= "    C2_QUANT," + CRLF
		cQuery+= "    C2_QUJE," + CRLF
		cQuery+= "    B1_UM," + CRLF
		cQuery+= "    B1_TIPO," + CRLF
		cQuery+= "    B1_GRUPO" + CRLF
		cQuery+= "FROM " + CRLF
		cQuery+= "    "+ RetSqlName("SC2") +" C2 (NOLOCK)" + CRLF
		cQuery+= "INNER JOIN " + CRLF
		cQuery+= "    "+ RetSqlName("SB1") +" B1 (NOLOCK) ON B1_COD = C2_PRODUTO " + CRLF
		cQuery+= "    AND B1.D_E_L_E_T_ = ''" + CRLF
		cQuery+= "WHERE " + CRLF
		cQuery+= "    C2_FILIAL = '"+ xFilial("SC2") +"'" + CRLF
		cQuery+= "    AND C2.D_E_L_E_T_ = ''" + CRLF
		cQuery+= "    AND C2_DATRF = ''" + CRLF
		cQuery+= "    AND C2_DATPRF <= '"+ DToS(_dtVencto) +"'" + CRLF
		If lProlinhas
			cQuery+= "    AND LEFT(C2_NUM,1)<>'R'" + CRLF
		Endif
		cQuery+= "ORDER BY C2_DATPRF"

		//MemoWrite("C:\temp\EST051.TXT", cQuery)
		TCQuery cQuery NEW ALIAS "TRB"
		TRB->(Dbgotop())
	
		// Gerar Excel  
	   	cWorkSheet := SM0->M0_FILIAL
	   	cTable     := cTitulo
	   
	   	oExcel:AddWorkSheet( cWorkSheet )
	   	oExcel:AddTable( cWorkSheet, cTable )	
	
		//Filial	DT Entrega	 Dias 	Ped Compra	Produto			Descrição					Tipo	Un.Med 	 Quantidade 	Já Entregue  	Fornecedor 	Loja	 Nome 	 Telefone 	 Contato 	Cotação	Solic.Compra	 Residuo (S/N)? 
		//03		15/10/2013	 	-6 	015363		PMPLBCM12000003	TRANSFORMADOR 220/20,4KW    MT	 	UN 	 	 2 	 			-   			000702	01	POLYTEK MAKINA DISTICARET LIMITED SIRKET	XX-XXXX-XXXX                                      	               	      	005472	 
		oExcel:AddColumn( cWorkSheet, cTable , "Filial"    			, 2,1)
		oExcel:AddColumn( cWorkSheet, cTable , "OP"					, 2,1)
		oExcel:AddColumn( cWorkSheet, cTable , "Emissão"			, 2,3)
		oExcel:AddColumn( cWorkSheet, cTable , "Dt. Prev. Início"	, 2,3)
		oExcel:AddColumn( cWorkSheet, cTable , "Dt. Prev. Fim"		, 2,3)
		oExcel:AddColumn( cWorkSheet, cTable , "Produto"			, 2,1)
		oExcel:AddColumn( cWorkSheet, cTable , "Qtde. Total"		, 2,2)
		oExcel:AddColumn( cWorkSheet, cTable , "Qtde. Já Entregue"	, 3,2)
		oExcel:AddColumn( cWorkSheet, cTable , "U.M."				, 2,1)
		oExcel:AddColumn( cWorkSheet, cTable , "Tipo"				, 2,1)
		oExcel:AddColumn( cWorkSheet, cTable , "Grupo"				, 2,1)
          
	   
		While TRB->(!Eof())
			If !lRotAuto   		
				IncProc()
			Endif      
			oExcel:AddRow( cWorkSheet, cTable, {TRB->C2_FILIAL,; 
													TRB->C2NUM,;
													SToD(TRB->C2_EMISSAO),; 
													SToD(TRB->C2_DATPRI),; 
													SToD(TRB->C2_DATPRF),; 
													TRB->C2_PRODUTO,; 
													TRB->C2_QUANT,;
													TRB->C2_QUJE,;
													TRB->B1_UM,;
													TRB->B1_TIPO,;
													TRB->B1_GRUPO})
		
			nReg++
			TRB->(dbSkip())
		End                            
		TRB->(dbCloseArea())            
	
		oExcel:Activate()
		cArq := CriaTrab( NIL, .F. ) + ".xls"
		//LjMsgRun( "Gerando o arquivo, aguarde...", cTitulo, {|| oExcel:GetXMLFile( cArq ) } )    
		oExcel:GetXMLFile( cArq )
		If !lRotAuto
			If __CopyFile( cArq, cDir + cArq )  //cDirTmp
				oExcel := MsExcel():New()
				oExcel:WorkBooks:Open( cDir + cArq ) //cDirTmp
				oExcel:SetVisible(.T.)
				oExcel:Destroy()
			//	MsgInfo( "Arquivo " + cArq + " gerado com sucesso no diretório " + cDir )
			Else
				MsgInfo( "Arquivo não copiado para temporário do usuário." )
			Endif
	    Endif

		// Enviar e-mail     
		_cFrom 	:= GetMV("MV_RELACNT")
		_cTo   	:="adelson.barbosa@prolinhas.com.br;Thayllive.pinheiro@prolinhas.com.br;giselly.santos@prolinhas.com.br;marlon.santos@prolinhas.com.br;lucas.oliveira@prolinhas.com.br"//GetMV("ZZ_MAILFLW")   //pablo coloca o seu email...
		_cCc    := "ana.jaize@prolinhas.com.br"
		_cCco   := "" 
		_cTitle	:= "[EST051] "+DTOC(date())+" "+Alltrim(SM0->M0_CODIGO)+" "+Alltrim(SM0->M0_CODFIL)+" - "+cTitulo+" - Tipo: "+If(x==1,"Somente MS","Outros, exceto MS e MP")
		_cFiles := cDir+cArq 
		_cType 	:= .T.
		_cBody  :="<html>"
		_cBody  +="<head><title> </title>"
		_cBody  +="<style type='text/css'>"
		_cBody  +="body{font-family:verdana;margin:0px;width:98%;height:auto;padding:1%;}"
		_cBody  +="</style>"
		_cBody  +="</head>"
		_cBody  +="<body>"
		_cBody  +="<p><b>Prezado(a), segue anexo: </b>"
		_cBody  +="<p><b> Relatório..: </b>"+"EST051 - Relação Ordens de Produção vencendo em até 5 dias"
		_cBody  +="<p><b> Empresa....: </b>"+Alltrim(SM0->M0_NOMECOM)
		_cBody  +="<p><b> Filial.....: </b>"+Alltrim(SM0->M0_CODIGO)+" "+Alltrim(SM0->M0_CODFIL)+" - "+Alltrim(SM0->M0_FILIAL)
		_cBody  +="<p><b> Local......: </b>"+Alltrim(SM0->M0_CIDENT)+"/"+SM0->M0_ESTENT
		_cBody  +="<br/br>"
		_cBody  +="<p><b> Emitido em.....: </b>"+DTOC(date()) +" "+Time()
		_cBody  +="</body>"
		_cBody  +="</html>"+CRLF
	
		u_WFEmail(_cFrom, _cTo, _cCc, _cCco, _cTitle, _cBody, _cFiles, _cType, lRotAuto)
	Next

Return Nil
