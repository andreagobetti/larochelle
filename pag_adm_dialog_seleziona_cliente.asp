<!--#include virtual="/setup.asp" -->
<%
'Verifica chiusure 02_12_2015
documento=replace(request.form("documento"),"nuovo_","")
if session("iduser")=1 then response.write queryeform()

%>
<form id="seleziona-cliente">
	<input type="hidden" name="documento" value="<%=documento%>" id="documento">
	<input  name="seleziona_iduser" id="seleziona_iduser" value="" style="width: 300px;"/>
	<div class="ui-dialog-buttonset" style="float: right;">
		<button type="button" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only avanti"  id="">
			<span class="ui-button-text">Avanti</span>
		</button>
	</div>
	<%if documento="ordine" then %>
	<div class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix" >
		<div class="ui-dialog-buttonset" style="float: left;">
			<button type="button" class="ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only avanti"  id="nuovo" >
				<span class="ui-button-text">Nuovo cliente</span>
			</button>
		</div>
	</div>
	<%end if %>
</form>


<script>
$(document).ready(function() {
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
	
	
	$('#seleziona_iduser').on("select2-selecting", function(e) {
		
		  $("#nuovo").hide();
    });
	$('#seleziona_iduser').on("select2-removed", function(e) {
		  $("#nuovo").show();
    });
	
	



    $(".avanti").click(function(e) {
        var url = '';
        var documento = $("#documento").val();

        if ($(this).attr("id") == "nuovo") {
            iduser = '';

        } else {
            iduser = $("#seleziona_iduser").val();

        }
        console.log("iduserrrr:" + iduser);

        switch (documento) {

            case 'ordine':
                url = 'pag_adm_dialog_ordine.asp';
                break;

            case 'ddt':
                url = 'pag_adm_dialog_creafattura.asp';
                var ddt_o_fattura = "D";
                break;
            case 'notacredito':
                url = 'pag_adm_dialog_ordine.asp';
                break;
            case 'sostituzione':
                url = 'pag_adm_dialog_ordine.asp';
                break;
            case 'preventivo':
                url = 'pag_adm_dialog_ordine.asp';
                break;
				
                
            case 'fattura':
            	url='pag_adm_dialog_ordine.asp';
                break;
            case 'ordine_fornitore':
            	url='pag_adm_dialog_ordine.asp';
                break;
        }


        $.ajax({
            url: url,
            type: "post",
            cache: false,
            //dataType: 'json',
            data: {
                //Provare con $("#seleziona-cliente").serialize();
                iduser: iduser,
                documento: documento,
                ddt_o_fattura: ddt_o_fattura
            },
            success: function(data) {
                $("#dialog").html(data);
                $("#dialog").dialog({
                    autoOpen: true,
                    modal: true,
                    resizable: true,
                    width: 600,
                    height: 'auto'
                    //title: "Crea ddt"
                });
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina', '', {
                    timeOut: 0
                });

                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in pulsante_crea_ddt_fattura()", txt);


            }
        });




    });



});
</script>