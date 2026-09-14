<!--#include virtual="/setup.asp" -->
<%
'Verifica chiusure 02_12_2015
iduser=request("iduser")
'Cerco di determinare il cliente principale
azienda=conn.execute("select azienda from utenti u inner join utenti_intestazioni i on u.idintestazione = i.id where u.iduser ="&iduser)(0)
sql="select * from utenti u inner join utenti_intestazioni i on u.idintestazione = i.id where u.iduser<>"&iduser&" and  i.azienda='"&azienda&"'"
set rs_cliente=conn.execute(sql)
if not rs_cliente.eof then
	response.write "iduser:"&rs_cliente("iduser")
	id=rs_cliente("iduser")
	testo=rs_cliente("azienda")&" id:"&id
end if

%>
<form id="seleziona-cliente" action="pag_adm_user.asp" method="POST">
	<input type="hidden" name="iduser" value="<%=iduser%>" >
	<input  name="iduser_primario" id="seleziona_iduser" value="" style="width: 300px;"/>
	<div class="ui-dialog-buttonset" style="float: right;">
		<button type="submit" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only avanti"  id="">
			<span class="ui-button-text">Conferma</span>
		</button>
	</div>
</form>


<script>
$(document).ready(function() {
    $('#seleziona_iduser').select2({
        placeholder: 'Seleziona cliente primario',
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