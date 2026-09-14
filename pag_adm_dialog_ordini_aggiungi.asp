<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
documento=request("documento")
nord=request("ndocumento")
dove=request("dove")
if dove="" then
	dover="sotto"
end if
iddett=request("iddett")
if iddett="" then iddett=0
ordine=request("ordine")
if ordine="" then ordine=0
idord=request("iddocumento")
idfor=request("idfor")

tabella=""
testo=""
gestisci_spettanze=true
call imposta_documento(documento)

if session("idadmin") = "" then call login()

oper=request("oper")
select case oper
	case "annulla"
		oper="view"
	case "new"
		oper="new"
	case "aggiungi"
		oper="add"
	case "modifica"
		oper="update"
	case else
		oper="new"
		if request("idpro_a")<>"" then oper="view"
end select
if oper="add" or oper="update" then
	'controlli
end if
tabellaid="idord"

if utente_andrea then
	response.write queryeform()
end if
%>
<%'response.write "ordine:"&ordine%>
<form id="form_ordine_aggiungi" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;" >
  <input type="hidden" name="tabella" value="<%=tabella%>" />
  <input type="hidden" name="dove" value="<%=dove%>" />
  <input type="hidden" name="ordine" value="<%=ordine%>" />
  <input type="hidden" name="iddett" value="<%=iddett%>" />
  <input type="hidden" name="idord" value="<%=idord%>" />
  <input type="hidden" name="nord" value="<%=nord%>" />
  <input type="hidden" name="documento" value="<%=documento%>">
	<% if gestisci_spettanze then
    	sql="select iduser from "&tabella&" where idord="&idord
		iduser=conn.execute(sql)(0)
     %>
		    Assegna a dipendente: <input type='hidden' name='iddip' id='iddip' style="width:350px;" tabindex="6" />
    <%end if %>

  <div id="aggiungiaordine">
  <ul>
    <li><a href="#aggiungiaordine-1">Articolo</a></li>
    <li><a href="#aggiungiaordine-2">Riga libera</a></li>
  </ul>
  <div id="aggiungiaordine-1">
  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
    <tr align="right" >
      <td  align="left" valign="bottom" style="border-bottom:1px solid; "><input type='hidden' name='idpro' id='idpro' style="width:350px;" tabindex="1" />
        <input type='hidden' name='idvara' id='vara' style="width:150px;" class="step1" tabindex="2" />
        <input type='hidden' name='idvarb' id='varb' style="width:150px;" class="step1" tabindex="3" />
        <br>
        <span style="float: right;">
         Quantit&agrave; 
        <input name="quantita" id="quantita" style="width:60px; text-align:right; " placeholder="quantità" tabindex="4" value="1"/><span id="quantitamagazzino"></span>
        <%if documento="ordine" or documento="fattura" or documento="ddt" or documento="preventivo" or documento="sostituzione" then %>
        Prezzo <input name="prezzo" id="prezzo"  style="width:60px; text-align:right; " tabindex="5">
       Sconto <input name="sconto" id="sconto"  style="width:60px; text-align:right; " tabindex="5">
        <%end if %>
       </span>
       <br>
       <%if documento="ordine" or documento="sostituzione" then

        %>
       Nota
	   <input type="text" name="nota" id="nota" value="" style="width: 98%;">
	   <%end if %>
       <p>
        <span style="float: right;">
        <input type="button" name="ordina" id="ordina" value="Aggiungi articolo"  disabled="disabled" tabindex="7" />
        </span>
        </p>
        </td>
    </tr>
  </table>
  </div>
  <div id="aggiungiaordine-2">
  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
    <tr align="right">
      <td  align="left" style="border-bottom:1px solid;"><textarea name="articolo_ordine" id="articolo_ordine" style="width:99%; " placeholder="Descrizione" rows="1"></textarea><br>
	  <span style="float: right; margin-top:5px;">
	  <select id="um">
		  <option value="">Nessuna u.m.</option>
		  <%
			  
			  if len(elencoum)>0 then
				  um=split(elencoum,VbCrLf)
			   	for i=0 to ubound(um)
				%>
          <option value="<%=um(i)%>" ><%=um(i)%></option>
          <%
				next
			end if	
		  %>
	  </select>
        <input name="quantita_1" type="text" id="quantita_1" size="8" placeholder="quantità" style="text-align:right;" >
        <input name="importo_1" type="text" id="importo_1" size="8" placeholder="importo" style="text-align:right;" title="-1 nasconde quantit&agrave;">
		</span><br>
        <p>
        <span style="float: right;"><input type="button" value="Aggiungi voce libera" id="aggiungi_voce_libera"></span>
        </p>
         </td>
    </tr>
    </table>
	</div>
	</div>
<%if esito<>"" then%>
  <br />
  <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
  <%if esiste<>"" then%>
    <tr align="right" bgcolor="#FF3300">
      <td align="left" style="border-bottom:1px solid;"><%=esito%><br />
      Articolo gi&agrave; presente nell'ordine</td>
    </tr>
  
  <%else %>
    <tr align="right" bgcolor="#33CC33">
      <td align="left" style="border-bottom:1px solid;"><%=esito%></td>
    </tr>
      <%end if%>
  </table>
  
    <%end if%>
  
	<div class="ui-widget" id="articolo_aggiunto" style="display: none;">
		<div class="ui-state-highlight alert-success ui-corner-all" style="margin-top: 20px; padding: 0 .7em;">
			<p><span class="ui-icon ui-icon-info" style="float: left; margin-right: .3em;"></span>
			<strong>OK</strong> <span id="testo"></span></p>
		</div>
	</div>
  
  
  
  
  <script type="text/javascript">
  $(document).ready(function () {
  	//$( "input" ).tooltip();
	$( "#aggiungiaordine" ).tabs();
	if($("#articolo_ordine")){
		$("#articolo_ordine").autogrow({onInitialize: true,speed: 50});
	}
  
  
	//Inizializzo select articolo
	var varianti = 0,
		variantea = false,
		varianteb = false,
		idpro = 0,
		idvara = 0,
		idvarb = 0,
		iddett = <%=iddett%>,
		dove ='<%=dove%>',
		ordine=<%=ordine%>;
		
		
	$("#ordina").click(function (e){
		e.preventDefault();
		console.log("ordina");
		var dati = {
			iddett:<%=iddett%>,
		    idpro:    	idpro,
		    idvara:		idvara,
		    idvarb:		idvarb,
		    quantita:	$("#quantita").val(),
		    dove: 		dove,
			ordine: 	ordine,
			aggiungo: "articolo",
			sconto: $("#sconto").val(),
			iddip: $("#iddip").val(),
			prezzo: $("#prezzo").val(),
			aggiungiarticolo: "si",
			nota: $("#nota").val()
			
		};
		
		if (aggiorna_tabella(dati)){
			//Aggiornamento eseguito
			mostra_avviso("Articolo aggiunto.");
			if(ordine>-1){
				ordine+=2;
			}
			$("#nota").val('');
		}
	});
	
	$("#aggiungi_voce_libera").click(function (e){
		e.preventDefault();
		console.log("aggiungi_voce_libera");
		if ($("#articolo_ordine").val() == "")  { 
			alert("Descrizione obbligatoria"); 
			$("#articolo_ordine").focus(); 
			return (false); 
		} 
		
		
		
		var dati = {
		    iddett: <%=iddett%>,
		    quantita:	$("#quantita_1").val(),
		    dove: 		dove,
			ordine: 	ordine,
			articolo_ordine: $("#articolo_ordine").val(),
			prezzo: 	$("#importo_1").val(),
			aggiungo: "vocelibera",
			iddip: $("#iddip").val(),
			um: $("#um").val(),
			aggiungiarticolo: "si"
		};
		if (aggiorna_tabella(dati)){
			mostra_avviso("Voce libera aggiunta.");
			//Aggiornamento eseguito
			if(ordine>-1){
				ordine+=2;
			}
		}
	});
	var timer;
	function mostra_avviso(testo){
		$("#articolo_aggiunto").hide();
		$("#testo").html(testo);
		$("#articolo_aggiunto").slideDown();
		clearTimeout(timer);
		timer= setTimeout(function(){
			  $("#articolo_aggiunto").slideUp();
			}, 2000);
		
	}
	
	
	
	
	
	
	
	
	$('#idpro').select2({
		dropdownCssClass: 'ui-dialog',
		placeholder: 'Seleziona articolo',
		minimumInputLength: 3,
		allowClear: true,
		ajax: {
			quietMillis: 150,
			url: "ajax_function.asp?select2=articoli_sel2",
			dataType: 'json',
			data: function (term, page) {
				return {
					term: term
				};
			},
			results: function (data) {
				return {
					results: data
				};
			}
		}
	});
	$("#idpro").on('change', function (e) {
		//console.log("valore idpro val:"+$("#idpro").select2('val'))
		//console.log("valore idpro data:"+$("#idpro").select2('data').id)
		console.log("idpro change");
		posizione_errore="#idpro_change";
		var data = $("#idpro").select2('data');
		$.ajaxSetup({
			cache: false
		});
		idvara = 0,
		idvarb = 0;
		varianti=0;
		variantea=false;
		varianteb=false;
		$("#quantitamagazzino").html( "" );
		posizione_errore="#idpro_change: P1";
		if (typeof data == "undefined") {
		   return true;
		}
		posizione_errore="#idpro_change: P1.1";
		if (!data){
			return true;
		}
		posizione_errore="#idpro_change: P1.2";
		idpro=data.id;
		posizione_errore="#idpro_change: P1.5";
		$("#prezzo").val("");
		//quantitamagazzino();
		$.getJSON("ajax_function.asp?select2=varianti&term=" + data.id)
			.done(function (json) {
				//console.log(json);
				//Verifico se ha varianti
				posizione_errore="#idpro_change: P2";
				if (json.var_a == "true") {
					varianti += 1;
					variantea = true;
				}
				if (json.var_b == "true") {
					varianti += 1;
					varianteb = true;
				}
				verifica();
				if (json.var_a == "false" && json.var_b == "false") {
					quantita();
				}
				console.log()
				if (json.var_a == "true") {
					posizione_errore="#idpro_change: P3";
					$.getJSON("ajax_function.asp?select2=vara&prezzo=si&term=" + data.id)
						.done(function (json) {
							//Inizializzo select variante A
							$('#vara').select2({
								dropdownCssClass: 'ui-dialog',
								data: json,
								placeholder: 'Variante A'
							});
							$("#vara").select2("val", "");
							
							quantita();
							$("#vara").on('change', function (e) {
								idvara=$("#vara").val();
								console.log("vara change:"+idvara);
								quantitamagazzino();
								verifica();
							});
						});
				} else {
					$('#vara').select2("destroy");
				}
				if (json.var_b == "true") {
					posizione_errore="#idpro_change: P4";
					$.getJSON("ajax_function.asp?select2=varb&term=" + data.id)
						.done(function (json) {
							//Inizializzo select variante A
							$('#varb').select2({
								dropdownCssClass: 'ui-dialog',
								data: json,
								placeholder: 'Variante B'
							});
							$("#varb").select2("val", "");
							quantita();
							$("#varb").on('change', function (e) {
								console.log("varb change");
								idvarb=$("#varb").val();
								//quantitamagazzino();
								verifica();
							});
						});
				} else {
					$('#varb').select2("destroy");
				}
			})
			.fail(function (jqxhr, textStatus, error) {
				var err = textStatus + ", " + error;
				console.log("Request Failed: " + err);
			});
	});
	function verifica() {
		//console.log("Inizio verifica variantea:"+variantea +" idvara:"+ idvara);
		posizione_errore="function verifica";
		var verifica = true;
		
		if (variantea && idvara == 0) {
				verifica = false;
		}
		if (varianteb && idvarb == 0) {
				verifica = false;
		}
		if (verifica){
			quantitamagazzino();
		}
		//console.log("Verifica:"+verifica)
		if ( $("#quantita").val() == "") {
			verifica = false;
			//console.log("Verifica quantita false");
		}
		
		//console.log("Fine verifica: "+verifica);
		if (verifica) {
			$("#ordina").removeAttr('disabled');
			//$("#ordina").focus();
			quantita();
		} else {
			$("#ordina").attr('disabled', 'disabled');
		}
	}
	function quantitamagazzino()
	{
		posizione_errore="function quantitamagazzino";
		if (! idvara) {
			idvara=0
		}
		if (! idvarb) {
			idvarb=0
		}
		console.log("idpro:"+idpro+" idvara:"+idvara+" idvarb:"+idvarb);
		$.ajaxSetup({ cache: false });
		$.ajax({
			url     : "ajax_function.asp?oper=magazzino&idpro="+idpro+"&idvara="+idvara+"&idvarb="+idvarb,
			type    : "post",
			dataType: "json",
			//data	: ,
			success: function(data){
				//console.log(data.success);
				//alert(data.Message);
					
				if (data.success){
					//$("#quantitamagazzino").html( data.quantita_magazzino );
					$("#quantitamagazzino").html( " / "+data.quantita_magazzino );
					$("#prezzo").val( String(data.prezzo).replace(".", ",") );
					console.log("quantita_magazzino:"+data.quantita_magazzino);
				}
				else
				{
					$("#quantitamagazzino").html( "" );
					//$("#quantitamagazzino").html( "errore");
				}
			}
			,error: function(xhr, textStatus, error){
				console.log("Errore ricerca quantita");
				invia_errore("Errore in quantitamagazzino()",xhr.statusText+xhr.responseText+textStatus+error );
			      
				  }
		});
	
	}
	
	function quantita() {
	
		$("#quantita").removeAttr("disabled");
		$("#quantita").on('change', function (e) {
			verifica();
		});
		$("#ordina").show()
	}
	
	<% if gestisci_spettanze then
	
	
	
	
	 %>
	
	
	 	var iduser=<%=iduser%>;
		$('#iddip').select2({
		dropdownCssClass: 'ui-dialog',
		placeholder: 'Seleziona un dipendente',
		minimumInputLength: 0,
		allowClear: true,
		ajax: {
			quietMillis: 150,
			url: "ajax_function.asp?select2=dipendenti",
			dataType: 'json',
			data: function (term, page) {
				return {
					term: term
					,iduser: iduser
				};
			},
			results: function (data) {
				return {
					results: data
				};
			}
		}
	});
	<%end if %>
	
	
});
</script>
</form>
<%
call connclose()
call CheckConnChiusa()
%>