<%
' =====================================================================================================================
' DESCRI��O	: Fun��es padr�o para todos os projetos
'	P�gina functions.asp
'
'	DATA: 31/10/2006 09:00			    AUTOR : Wilson Roberto P. J�nior
' =====================================================================================================================
' MANUTEN��ES
'	Descri��o	: -
'	Data		:						Autor :
' =====================================================================================================================
%>

<%
	' Declara��o de vari�veis
	session.LCID						= 1046
	session.Timeout						= 1200
	server.ScriptTimeout				= 12000

	' ***** DESENVOLVIMENTO *****
	Public Const cstr_conexao			= "Driver={SQL Server};Server=LOCALHOST;Database=db_entrevista;UID=desenvolvimento;PWD=newi;"
   
	vstr_tituloSite						= "Portal do Cliente"
    vstr_statusSite						= "Portal do Cliente"
	vstr_emailSite                      = "no-reply@newi.com.br"

	' Enumera��o das constantes que especifica como um comando argumento poderia ser interpretado.
	Public Const cint_adCmdStoredProc	= 4

	' Vari�vel objeto que refer�ncia o objeto de comunica��o com o banco de dados.
	Public vobj_conexao

	vstr_caminho		= request.ServerVariables("URL")
	vint_posicao		= InstrRev(vstr_caminho, "/", len(vstr_caminho), 0)
	vstr_nomePagina	    = right(vstr_caminho,(len(vstr_caminho)-vint_posicao))

    'grava na variave a pagina para bater no perfil do usuario para liberar acesso
    if instr(vstr_pasta,"admin") > 0 then
        vstr_paginaPerfil = "admin/" & vstr_nomePagina
    else
        vstr_paginaPerfil = vstr_nomePagina
    end if

%>

<%

' -------------------------------------------------------------------------------
' Nome Fun��o	:	fcn_abrirConexao()
' Par�metros	:	Nenhum
' Retorno		:	Disponibilidade de conex�o com a v�riavel vobj_conexao
' Descri��o		:	Abre a conex�o com o Banco de Dados
' -------------------------------------------------------------------------------
Public Function fcn_abrirConexao()

	Set vobj_conexao = Server.CreateObject("ADODB.Connection")

	vobj_conexao.Open cstr_conexao

End Function

' -------------------------------------------------------------------------------
' Nome Fun��o	:	fcn_fecharConexao()
' Par�metros	:	Nenhum
' Retorno		:	Disponibilidade de conex�o com a v�riavel vobj_conexao
' Descri��o		:	Procedimento desenvolvido para fechar a conex�o do
'					objeto vobj_conexao
' -------------------------------------------------------------------------------
Public Sub fcn_fecharConexao()
	If Not vobj_conexao Is Nothing Then

		If vobj_conexao.State <> adStateClosed Then
			vobj_conexao.Close
		End If

		Set vobj_conexao = Nothing
	End If
End Sub

' -------------------------------------------------------------------------------
' Nome Fun��o	:	fcn_limparString(pstr_string)
' Par�metros	:	pstr_string - qualquer string
' Retorno		:	string sem aspas e caracteres proibidos
' Descri��o		:	Fun��o limpa a string passada se ela
' -------------------------------------------------------------------------------
Function fcn_limparString(pstr_string)

	'Declara v�riavel para limpar string
	Dim vstr_string

	'Tira espa�os em brancos
	vstr_string			= Trim(pstr_string)

	If len(vstr_string)>0 then
		vstr_string			= Replace(Replace(Replace(vstr_string,"'",""),"&",""),"""","")
	end if

	If vstr_string		= "" OR vstr_string	= " " Then
		vstr_string		= empty
	End If

	'Retornando valor
	fcn_limparString	= vstr_string

End Function
%>
