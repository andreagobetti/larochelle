<!--#include virtual="/setup.asp" -->
<%
'Verifica chiusure 02_12_2015
iddocumento=request.querystring("id_documento")
documento=request.querystring("documento")

if request.form("seleziona_iduser")<>"" then
        documento=request.form("documento")
        iddocumento=request.form("iddocumento")
        if documento="ordine" then
            tabella="ordini"
            idfiled="idord"
        end if
        if documento="ddt" then
            tabella="ordini"
            idfiled="idord"
        end if
        if documento="fattura" then
            tabella="fatture"
            idfiled="IDfat"
        end if

        iduser=request.form("seleziona_iduser")


        sql="SELECT utenti.* FROM utenti where iduser="&iduser
        set rs_cliente=conn.execute(sql)

        sql="SELECT "&tabella&".* FROM "&tabella&" where "&idfiled&"="&iddocumento
		Set rs = Server.CreateObject("ADODB.Recordset")
		response.write sql
		rs.Open sql, conn, 3, 3
		rs("iduser")=iduser
		rs("idintestazione")=rs_cliente("idintestazione")
		rs.update
		rs.close
        set rs=Nothing


        if documento="ordine" then
            response.Redirect "/pag_adm_ordini.asp"
        end if
        if documento="ddt" then
            response.Redirect "/pag_adm_ddt.asp"
        end if
        if documento="fattura" then
            response.Redirect "/pag_adm_fatture.asp"
        end if


end if

if session("iduser")=1 then
 response.write("documento:"&documento)

end if
%>

<form action="pag_admin_cambia_intestazione.asp" method="post">
	<input type="hidden" name="documento" value="<%=documento%>">

	<input type="hidden" name="iddocumento" value="<%=iddocumento%>" id="documento">
	<input  name="seleziona_iduser" id="seleziona_iduser" value="" style="width: 300px;"/>
	<input type="submit" name="Avanti" value="Conferma">
</form>
<link rel="stylesheet" type="text/css" href="/Jquery/css/select2.css">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
<script src="/Jquery/js/select2/select2.js"></script>
<script src="/Jquery/js/select2/select2_locale_it.js"></script>
<script>

$(function() {
    $('#seleziona_iduser').select2({
        placeholder: 'Seleziona cliente',
        minimumInputLength: 1,
		allowClear: true,
        width: 'resolve',
        ajax: {
            quietMillis: 150,
            url: 'ajax_function.asp?select2=utenti_sel2',
            dataType: 'json',
            data: function(term, page) {
                return {
                    term: term
                };
            },
            results: function(data) {
                return {
                    results: data
                };
            }
        }
    });



});
</script>
