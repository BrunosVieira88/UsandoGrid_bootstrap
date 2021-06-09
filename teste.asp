<!--#include file="functions.asp"-->

<!DOCTYPE html>
<html lang="pt">
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" type="text/css" href="style.css">
    <meta name="description" content="">
    <meta name="author" content="">
        
</head>
    <title><%=vstr_tituloSite%></title>

    <link rel="shortcut icon" href="<%=vstr_local%>imagens/icos/favicon.ico" />

</head>

<%
	if len(request("hdn_operacao")) > 0 then

		if request("hdn_operacao") = 1 then
                
            'Abre a conexao
			Call fcn_abrirConexao()

			'Cria objeto para Proc
			Set vobj_command							        = Server.CreateObject("ADODB.Command")
			Set vobj_command.ActiveConnection			        = vobj_conexao

			'Iniciando Proc
			vobj_command.CommandText					        = "NEW_SP_usuario"
			vobj_command.CommandType					        = cint_adCmdStoredProc
			vobj_command.Parameters.Refresh

			'Passando os parametros para PROC
			vobj_command.Parameters("@vstr_tipoOper")	        = "INS"
			vobj_command.Parameters("@vint_numOper")	        = 1
            vobj_command.Parameters("@ds_nome")		            = fcn_limparString(request("txt_nome"))
            vobj_command.Parameters("@ds_nascimento")           = fcn_limparString(request("txt_nascimento"))
            vobj_command.Parameters("@ds_genero")               = fcn_limparString(request("txt_genero"))
			vobj_command.Parameters("@ds_email")    	        = fcn_limparString(request("txt_email"))
            vobj_command.Parameters("@ds_telefone")    	        = fcn_limparString(request("txt_telefone"))
            vobj_command.Parameters("@ds_cidade")    	        = fcn_limparString(request("txt_cidade"))
            vobj_command.Parameters("@ds_estado")    	        = fcn_limparString(request("txt_estado"))
            vobj_command.Parameters("@fl_status")    	        = 1
            vobj_command.Parameters("@id_usuarioInclusao")    	= 1

			'Executa procedure
			vobj_command.Execute

            'Limpa Objeto
			Set vobj_command = Nothing

			'Fecha a conexao
			Call fcn_fecharConexao()

            response.Redirect "teste.asp?msg=Usuario_gravado_com_sucesso"

		end if

	end if
%>	

<body id="page-top" style="background-color:#f0effe">
    <form id="frm_newi" name="frm_newi" action="" method="post">
    <input type="hidden" name="hdn_operacao"  id="hdn_operacao" />
    
		<div class="text-center mb-5" id="text_texto_principal">
            <h1 class="h4 text-gray-900 mb-4">TESTE CANDIDATO NEWI</h1>
            <hr>
        </div>

        <%
            if len(request("msg")) > 0 then
        %>
                <script>
                    //alert('<%=request("msg")%>');
                </script>

                <div style="text-align: center;color:red;font-weight: bold;padding: 10px;">
                    <%=request("msg")%>
                </div>
        <%
            end if
        %>

        <section class="col-md-12 "  id="ds_formulario" >
                <div class="row justify-content-center mb-3 " >
                    <div class="col-auto">
                        <i class="fas fa-user"  style="padding-top:10px"></i>
                    </div>
                    <div class="col-8" >
                      Nome: <input type="text" autocomplete="off" maxlength="200" name="txt_nome" id="txt_nome" class="form-control form-control-user"  placeholder="Nome" required />
                    </div>
                                            
                </div>

        		<br />

                <div class="row justify-content-center mb-3 ">
                    <div class="col-auto">
                        <i class="fas fa-user"  style="padding-top:10px"></i>
                    </div>
                    <div class="col-8" >
                      Nascimento: <input type="date"  name="txt_nascimento" id="txt_nascimento" class="form-control form-control-user"  required />
                    </div>
                                            
                </div>

                <br />


                <div class="row justify-content-center mb-3 " >
                    <div class="col-auto">
                        <i class="fas fa-user"  style="padding-top:10px"></i>
                    </div>
                    <div class="col-8" >
                        Masculino: <input type="radio" style="margin-right:25px;" autocomplete="off"  name="txt_genero" id="txt_masculino" class="form-control form-control-user" value="M"  checked/>
                        Feminino: <input type="radio"  autocomplete="off"  name="txt_genero" id="txt_feminino" class="form-control form-control-user" value="F" placeholder="Nome"/>
                    </div>
                                            
                </div>

                <br />

                <div class="row justify-content-center mb-3">
                    <div class="col-auto">
                        <i class="fas fa-asterisk" style="padding-top:10px"></i>
                    </div>
                    <div class="col-8">
                        E-mail: <input type="text" autocomplete="off" maxlength="200" name="txt_email" id="txt_email" class=" form-control form-control-user" placeholder="E-mail" required/>
                    </div>
                </div>

                <br />

                <div class="row justify-content-center mb-3">
                    <div class="col-auto">
                        <i class="fas fa-asterisk" style="padding-top:10px"></i>
                    </div>
                    <div class="col-8">
                        Tel:<input type="tel" maxlength="11" style="margin-left: 45px;" autocomplete="off" maxlength="200" name="txt_telefone" id="txt_telefone" class=" form-control form-control-user" placeholder="Telefone" required/>
        		    </div>
                </div>

        		<br />

              <div class="row justify-content-center mb-3 " >
                 <div class="col-auto">
                        <i class="fas fa-user"  style="padding-top:10px"></i>
                    </div>
                    <div class="col-8">
                        Cidade: <input type="text" autocomplete="off" maxlength="200" name="txt_cidade" id="txt_cidade" class="form-control form-control-user" placeholder="cidade" required />
                    </div>
              </div>
                	<br />


        <div class="row justify-content-center mb-3 " >
                 <div class="col-auto">
                        <i class="fas fa-user"  style="padding-top:10px"></i>
                    </div>
                    <div class="col-8">
                        Estado: <input type="text" autocomplete="off" maxlength="200" name="txt_estado" id="txt_estado" class="form-control form-control-user" placeholder="estado" />
         </div>           </div>
                
                    <br />

          

                <div class="row justify-content-center mb-5">
                    <div class="col-xl-5 col-lg-12 text-center ">
                        <input type="button" id="btn"class="btn btn-default background-logo text-white w-100" value="Gravar" onclick="javascript:fcn_gravar();" />
                    </div>
                </div>    
        </section>


            </form>
       
</body>
</html>

<script type="text/javascript">
    
    function fcn_gravar() {


        if (document.frm_newi.txt_nome.value == "") {
            alert("Preencha o Nome!");
            document.frm_newi.txt_nome.focus();
	        return false;
	    }
        if (document.frm_newi.txt_nascimento.value == "") {
            alert("Preencha o nascimento!");
            document.frm_newi.txt_nascimento.focus();
            return false;
        }
        if (document.frm_newi.txt_genero.value == "") {
            alert("Preencha o Genero!");
            document.frm_newi.txt_genero.focus();
            return false;
        }
        if (document.frm_newi.txt_genero.value == "") {
            alert("Preencha o Genero!");
            document.frm_newi.txt_genero.focus();
            return false;
        }
          
        if (document.frm_newi.txt_email.value == "") {
            alert("Preencha o E-mail!");
            document.frm_newi.txt_email.focus();
            return false;
        }

        if (document.frm_newi.txt_telefone.value == "") {
            alert("Preencha o Telefone!");
            document.frm_newi.txt_telefone.focus();
            return false;
        }

        if (document.frm_newi.txt_cidade.value == "") {
            alert("Preencha a cidade!");
            document.frm_newi.txt_cidade.focus();
            return false;
        }
        if (document.frm_newi.txt_estado.value == "") {
            alert("Preencha o Estado!");
            document.frm_newi.txt_estado.focus();
            return false;
        }

        if (confirm("Deseja realmente gravar?")) {
            document.frm_newi.hdn_operacao.value = 1;
            document.frm_newi.submit();
            return false;
        }

	}

</script>