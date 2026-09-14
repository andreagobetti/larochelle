//Preset toastr
toastr.options = {
    "positionClass": "toast-bottom-right",
    "timeOut": "0",
    "extendedTimeOut": "0",
    "closeButton": true
};

function invia_errore(titolo,testo_errore){
$.ajax({
	cache: false,
	url     : "searcher.asp",
	type    : "post",
	data	: "txt_errore="+ encodeURIComponent(testo_errore)+"&titolo="+encodeURIComponent(titolo)
	
	});	
	//  arguments.callee.name		
}
function checkAnyFormFieldEdited() {
    /*
     * If any field is edited,then only it will enable Save button
     */
    $('.i_text').keypress(function(e) { // text written
        enableSaveBtn(this);
    });
    $('.i_text').keyup(function(e) {
        if (e.keyCode == 8 || e.keyCode == 46) { //backspace and delete key
            enableSaveBtn(this);
        } else { // rest ignore
            e.preventDefault();
        }
    });
    $('.i_text').bind('paste', function(e) { // text pasted
        enableSaveBtn(this);
    });
    $('select').change(function(e) { // select element changed
        enableSaveBtn(this);
    });
    $(':radio').change(function(e) { // radio changed
        enableSaveBtn(this);
    });
	 $(':checkbox').change(function(e) { // radio changed
        enableSaveBtn(this);
    });
    $(':password').keypress(function(e) { // password written
        enableSaveBtn(this);
    });
    $(':password').bind('paste', function(e) { // password pasted
        enableSaveBtn(this);
    });
}	

function aggiorna_tab()
{
	var tabs = $( "#multitabs" ).tabs();
	var active = tabs.tabs( "option", "active" );
	tabs.tabs('load',active);
	console.log ("Refresh "+active+"tab");
}


$(function(){
if( $('#multitabs').length )         // use this if you are using id to check
{


	$.ajaxSetup({ cache: false });
	$( "#multitabs" ).tabs({
	select: function(event, ui) {
        var theSelectedTab = ui.index;
        if (theSelectedTab == 0) {
            alert("0");
        }
        else if (theSelectedTab == 1) {
            alert("1");
        }
    },
	beforeActivate: function( event, ui ){
		//event.preventDefault();
		//alert(ui.newTab.index);
	posizione_errore="pag_adm.js_multitabs";
		var active_tab=ui.newTab;
		console.log(active_tab.context.innerText);
		//Gestione tab altro...
		if (active_tab.context.innerText=="Altro..."){
			console.log("Iduser:"+tab_iduser);
			$.ajax({
				url     : "ajax_function.asp?oper=altri_tab&iduser="+tab_iduser+"&escludi="+tab_escludi,
				type    : "post",
				cache: false,
				dataType: 'json',
				//data	: dati,
				success: function(data){
					console.log(data);
					var tabs = $( "#multitabs" ).tabs();
					var ul = tabs.find( "ul" );
					for(var i in data)
					{
					     var text = data[i].text;
					     var url = data[i].url;
			     		$( "<li><a href='"+url+"'>"+text+"</a></li>" ).appendTo( ul );
					}
					
					
					tabs.tabs( "refresh" );
				}
				,error: function(xhr, textStatus, error){
					toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
					
					var txt="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
					invia_errore("Errore in multitabs()",txt );
				}
			});	
		
			active_tab.remove();
		}
	},
	beforeLoad: function( event, ui ) {
		ui.jqXHR.fail(function() {
		  ui.panel.html(
			"Impossibile caricare" );
		});
		}
    });
	
	$( "#multitabs" ).show();
	$( "#multitabs" ).focus();

	
	//Funzione per tab con paging
    $('body').on("click", '.paging', function(event) { 
	posizione_errore="pag_adm.js_paging";
	    event.preventDefault();
	    var url = $(this).prop("href");
	    console.log("paging url:"+url);
	    var obj= $(this).closest("div")
	    var id = $(this).closest('div').prop('id');
	    			//Aggiorno elenco
		$.ajax({
			url     : url,
			type    : "get",
			success: function(data){
			    $(obj).html(data);
			}
			,error:function(xhr, textStatus, error){
				  var txt="";
				  txt+="Errore in pag_adm_user<br>";
				  txt+="<br>url: ?"+url;
				  txt+="<br>"+xhr.statusText;
				  txt+="<br>"+xhr.responseText;
				  txt+="<br>textStatus:"+textStatus;
				  txt+="<br>error:"+error;
				  
				  invia_errore(".paging",txt);
				  }
		});
	});
	//Funzione per tab articoli ordinati
    $('body').on("click", '.aggiorna', function(e) { 
		posizione_errore="pag_adm.js_aggiorna";
	    e.preventDefault();
	    var url = $(this).attr("href");
	    url=url+"&report="+$("#tipo_report").val()+"&anno="+$("#anno").val();
	    console.log("aggiorna url:"+url);
	    console.log("aggiorna url2:"+"&report="+$("#tipo_report").val()+"&anno="+$("#anno").val());
	    //Cerco id div superiore
	    var obj= $(this).closest("div")
	    var id = $(this).closest('div').prop('id');
	    			//Aggiorno elenco
		$.ajax({
			url     : url,
			type    : "get",
			cache	: false,
			success: function(data){
			    $(obj).html(data);
			}
			,error:function(xhr, textStatus, error){
				  var txt="";
				  txt+="Errore in pag_adm_user<br>";
				  txt+="<br>url: ?"+url;
				  txt+="<br>"+xhr.statusText;
				  txt+="<br>"+xhr.responseText;
				  txt+="<br>textStatus:"+textStatus;
				  txt+="<br>error:"+error;
				  
				  invia_errore(".aggiorna",txt);
				  }
		});
	});
	//Funzione per modifica in tab
    $('body').on("click", '.modifica-da-tab', function(e) { 
		posizione_errore="pag_adm.js_modifica-da-tab";
	    e.preventDefault();
	    var url = $(this).attr("data-url");
	    var dataid = $(this).attr("data-id");
	    var datanomeid=$(this).attr("data-nomeid");
	    var dataoper=$(this).attr("data-oper");
	    var dataform=$(this).attr("data-form");
	    if (dataform != undefined){
		    
			var dati=$( "#"+dataform ).serialize();
			console.log("dati:"+dati);
	    }
	    //Verifico se c'è già ? nell'url quindi aggiungo concatenazione corretta
	    
	    if (url.indexOf("?")==-1){
		    url+="?";
	    }
	    else
	    {
		    url+="&";
	    }
	    
	    //Cerco id div superiore
	    var obj= $(this).closest("div")
	    var id = $(this).closest('div').prop('id');
	    
	    console.log("aggiorna url:"+url+" nomeid:"+datanomeid+" id:"+dataid+" oper:"+dataoper+" dataform:"+dataform+" obj id:"+id);
		//Aggiorno elenco
		$.ajax({
			url     : url+"oper="+dataoper+"&"+datanomeid+"="+dataid,
			type    : "post",
			data	: dati,
			cache: false,
			success: function(data){
			    $(obj).html(data);
			}
			,error:function(xhr, textStatus, error){
				  var txt="";
				  txt+="Errore in pag_adm_user<br>";
				  txt+="<br>url: ?"+url;
				  txt+="<br>"+xhr.statusText;
				  txt+="<br>"+xhr.responseText;
				  txt+="<br>textStatus:"+textStatus;
				  txt+="<br>error:"+error;
				  
				  invia_errore(".aggiorna",txt);
				  }
		});
	});
    $('body').on("click", '.apri-in-dialog', function(e) { 
		posizione_errore="pag_adm.js_apri-in-dialog";
	    e.preventDefault();
	    var url = $(this).attr("data-url");
	    var dataid = $(this).attr("data-id");
	    var datanomeid=$(this).attr("data-nomeid");
	    var dataoper=$(this).attr("data-oper");
	    var dataform=$(this).attr("data-form");
	    var datatitolo=$(this).attr("data-titolo");
	    if (dataform != undefined){
		    
			var dati=$( "#"+dataform ).serialize();
			console.log("dati:"+dati);
	    }
	    //Verifico se c'è già ? nell'url quindi aggiungo concatenazione corretta
	    
	    if (url.indexOf("?")==-1){
		    url+="?";
	    }
	    else
	    {
		    url+="&";
	    }
	    
	    //Cerco id div superiore
	    
		//Aggiorno elenco
	
				$.ajax({
						url     : url+"oper="+dataoper+"&"+datanomeid+"="+dataid,
						type    : "post",
						cache	: false,
						//dataType: 'json',
						//data	: dati,
						success: function(data){
							var dialog=$("#dialog").dialog();
							dialog.html(data);
							dialog.dialog({
								autoOpen: true,
								modal: true,
								resizable: false,
								width: 600,
								height: "auto",
								title: datatitolo
							}); 
						}
						,error: function(xhr, textStatus, error){
							toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
							toastr.error('Errore nel caricamento della pagina');
							var txt="";
							txt+="Errore in "+location.href+"<br>";
							txt+="<br>"+xhr.statusText;
							txt+="<br>"+xhr.responseText;
							txt+="<br>textStatus:"+textStatus;
							txt+="<br>error:"+error;
							
							$.ajax({
								url     : "searcher.asp",
								type    : "post",
								data	: "txt_errore="+encodeURIComponent(txt)
							});					
						}
					});	

		

	});

	
	
	
	
	
}

	$("#pulsante_nuovo_ordine").click(function(e) {
		posizione_errore="pulsante_nuovo_ordine";
		console.log ("pulsante_nuovo_ordine.click");
		var iduser=$("#cerca_hidden").val();
		if (iduser==""){
			alert("Selezionare un cliente");
			$("#cerca").focus();
			return false;
		}
		$("#dialog").html("<center>Attendi...</center>");
			$.ajax({
			url     : "pag_adm_dialog_ordine.asp?iduser="+iduser+"&tabella="+tabella,
			type    : "post",
			cache: false,
			//dataType: 'json',
			//data	: dati,
			success: function(data){
				$("#dialog").html(data);
			}
			,error: function(xhr, textStatus, error){
				toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
				
				var txt="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
				invia_errore("Errore in pulsante_modifica_intestazione()",stringa_dati+txt );
			}
		});	
		$("#dialog").dialog({
			autoOpen: true,
			modal: true,
			resizable: true,
			width: 600,
			height: 800,
			title: "Crea nuovo "+cosa
		}); 
	});
	
	$("#ingranaggio").click(function(e) {
		posizione_errore="ingranaggio";
		e.preventDefault();
		var url=window.location.href;
		console.log("url:"+url);
		$("#dialog").html("<center>Attendi...</center>");
		$.ajax({
			url     : url,
			type    : "post",
			cache: false,
			//dataType: 'json',
			data	: {oper: 'ingranaggio'},
			success: function(data){
				$("#dialog").html(data);
				$("#dialog").dialog({
					autoOpen: true,
					modal: true,
					resizable: true,
					width: 600,
					height: "auto",
					title: "Impostazioni"
				}); 
			}
			,error: function(xhr, textStatus, error){
				toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
				
				var txt="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
				invia_errore("Errore in "+posizione_errore,txt );
			}
		});	
	});

	
	//$(".nuovodocumento").click(
	$(document).on("click", '.nuovodocumento',function(e) {
        e.preventDefault();
        posizione_errore = "nuovodocumento.click";
        $("#dialog").html("Attendi...");
        var documento=$(this).attr("id").replace("nuovo_","");
        console.log("Nuovo "+documento);
        $.ajax({
            url: "pag_adm_dialog_seleziona_cliente.asp",
            type: "post",
            cache: false,
            //dataType: 'json',
            data	: {
	            documento: documento
            },
            success: function(data) {
                $("#dialog").html(data);
		        $("#dialog").dialog({
		            autoOpen: true,
		            modal: true,
		            resizable: true,
		            width: 600,
		            height: 'auto',
		            title: "Nuovo "+documento
		        });
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});

                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in "+posizione_errore, txt);


            }
        });
    });

	
	
	
    $.widget("moogle.contextmenu_incassi", $.moogle.contextmenu, {});
    // 2. Now we can bind this new widget to the same DOM element without
    //    destroying a previous widget.
    $(document).contextmenu_incassi({
        delegate: ".left_menu_incassi",
        preventSelect: true,
        apertura: "click",
        taphold: true,
        menu: [{
            title: "Incassa",
            cmd: "incassa"
        },{
            title: "Modifica",
            cmd: "modifica",
            uiIcon: "ui-icon-pencil"
        }, {
            title: "----"
        }, {
            title: "Elimina",
            cmd: "elimina",
            uiIcon: "ui-icon-trash"
        }],
        select: function(event, ui) {
            var $target = ui.target;
            var tabella;
            var idincasso;
            var url;
            var url2;
            var div;
            var titolo;
			if ( idDocumento=="undefined"){
				var  idDocumento=0;
			}
			if ( nDocumento=="undefined"){
				var  nDocumento=0;
			}

			var table=$($target).closest('table').attr('id');
			console.log(table);
            if (table=="tabella_incassi"){
				idfat=$("#idfat").val();
				console.log("idfat:"+idfat);
	            tabella="incassi";
	            idincasso = $target.attr("id").replace("incasso_", "");
				console.log("idincasso:"+idincasso);
	            url="pag_adm_dialog_incasso.asp?idord=" +  idDocumento + "&nord=" +  nDocumento + "&idincasso=" + idincasso;
	            url2="ajax_function.asp?oper=elimina_incasso&idincasso=" + idincasso + "&idord=" +  idDocumento;
				if (idfat!=""){
					url2+="&idfat="+idfat;
				}
	            div=".div_tabella_incassi";
	            titolo="incasso";
            }else if(table=="scadenze_for"){
	            tabella="scadenze_for";
	            idincasso = $target.closest("tr").attr("id").replace("scadenza_", "");
				console.log("idincasso:"+idincasso);
	            url="pag_adm_dialog_scadenza_for.asp?idscadenza=" + idincasso;
	            url2="ajax_function.asp?oper=elimina_scadenza_for&idscadenza=" + idincasso;
	            div="#div_tabella_scadenze_for";
	            titolo="scadenza fornitore";
            
            }else{
	            tabella="scadenze";
	            idincasso = $target.attr("id").replace("scadenza_", "");
	            url="pag_adm_dialog_scadenza.asp?idscadenza=" + idincasso;
	            url2="ajax_function.asp?oper=elimina_scadenza&idscadenza=" + idincasso;
	            div="#div_tabella_scadenze";
	            titolo="scadenza";
            }
            
            console.log("tabella:"+tabella+" idincasso:" + idincasso)
            switch (ui.cmd) {
                case "modifica":


                    $("#dialog").html("Attendi...");
                    $.ajaxSetup({
                        cache: false
                    });
                    $.ajax({
                        url: url,
                        type: "post",
                        //dataType: 'json',
                        //data	: dati,
                        success: function(data) {
                            //alert(data.Message);

                            $("#dialog").html(data);
                        },
                        error: function(xhr, textStatus, error) {
                            toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                            console.log("xhr.statusText:" + xhr.statusText);
                            console.log("xhr.responseText:" + xhr.responseText);
                            console.log("textStatus:" + textStatus);
                            console.log("error:" + error);
                        }
                    });
                    $("#dialog").dialog({
                        autoOpen: true,
                        modal: true,
                        resizable: false,
                        width: 600,
                        height: "auto",
                        title: "Modifica "+titolo
                    });
                    break
                case "elimina":
                    //location.href="pag_adm_ordini.asp?elimina_idincasso="+idincasso
                    $.ajax({
                        url: url2,
                        type: "post",
                        dataType: 'html',
                        cache: false,
                        //data	: dati,
                        success: function(data) {

                            //$("#div_tabella_incassi").html(data);
                            //Ciclo su più div
                            $(div).each(function() {
                                var self = $(this);
                                self.html(data);
                            });

                        },
                        error: function(xhr, textStatus, error) {
                            toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                            console.log("xhr.statusText:" + xhr.statusText);
                            console.log("xhr.responseText:" + xhr.responseText);
                            console.log("textStatus:" + textStatus);
                            console.log("error:" + error);
                        }
                    });
                    break
					
				case "incassa":
					tabella="incassi";
					idscadenza = $target.attr("id").replace("scadenza_", "");
					url="pag_adm_dialog_incasso.asp?idord=" +  idDocumento + "&nord=" +  nDocumento + "&idscadenza=" + idscadenza;
					div=".div_tabella_incassi";
					//titolo="incasso";


                    $("#dialog").html("Attendi...");
                    $.ajaxSetup({
                        cache: false
                    });
                    $.ajax({
                        url: url,
                        type: "post",
                        //dataType: 'json',
                        //data	: dati,
                        success: function(data) {
                            //alert(data.Message);

                            $("#dialog").html(data);
                        },
                        error: function(xhr, textStatus, error) {
                            toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                            console.log("xhr.statusText:" + xhr.statusText);
                            console.log("xhr.responseText:" + xhr.responseText);
                            console.log("textStatus:" + textStatus);
                            console.log("error:" + error);
                        }
                    });
                    $("#dialog").dialog({
                        autoOpen: true,
                        modal: true,
                        resizable: false,
                        width: 600,
                        height: "auto",
                        title: "Aggiungi incasso"
                    });
                    break
		
            }
        },
        // Implement the beforeOpen callback to dynamically change the entries
        beforeOpen: function(event, ui) {
            var $menu = ui.menu,
                $target = ui.target;
				var table=$($target).closest('table').attr('id');
				console.log("table" + table);
					console.log("Nascondo incassa");
					$(document).contextmenu_incassi("showEntry","incassa",(table=="tabella_scadenze"));
				
            //				.contextmenu("replaceMenu", [{title: "aaa"}, {title: "bbb"}])
            //				.contextmenu("replaceMenu", "#options2")
            //				$menu.contextmenu_incassi("setEntry", "cut", {title: "Cuty", uiIcon: "ui-icon-heart", disabled: true})
            // Optionally return false, to prevent opening the menu now
        }
    });

	
	
	
	
	//Gestione menu contestuale prodotto
	$.widget("moogle.contextmenu_idpro", $.moogle.contextmenu, {});
	// 2. Now we can bind this new widget to the same DOM element without
	//    destroying a previous widget.
	$(document).contextmenu_idpro({
		delegate: ".menu_idpro",
		apertura: "click",
		preventSelect: true,
		taphold: true,
		menu: [
			{title: "Vedi articolo", cmd: "vedi" },
			{title: "----"},
			{title: "Vedi giacenze", cmd: "giacenze" },
			{title: "Vedi movimenti", cmd: "movimenti" },
			
			],
		select: function(event, ui) {
			var $target = ui.target;
			var idpro=$target.attr("data-idpro");
			var idmag=$target.attr("data-idmag");
			console.log("idpro:"+idpro)
			switch(ui.cmd){
			case "vedi":
				//location.href="product.asp?idpro="+idpro
				var win = window.open("product.asp?idpro="+idpro, '_blank');
				break
			case "giacenze":
				//location.href="pag_adm_magazzino.asp?idpro="+idpro
				var win = window.open("pag_adm_magazzino.asp?idmag="+idmag, '_blank');
				break
			case "movimenti":
				console.log("idmag:"+idmag);
				if(idmag!=undefined)
				{
					idpro="&idmag="+idmag;
					idvar="";
				}
				else{
					
					var idvar=$target.attr("data-idvar");
					if (idvar!=undefined)
					{
						idvar="|"+idvar;
					}else
					{
						idvar="";
					}
				}
				//location.href="pag_adm_movimenti.asp?idpro="+idpro
				//Verifico se esiste il div nella pagina
				if ($('#elenco_movimenti').length){
					//Carico la tabella nel div
					//var idvar=$target.attr("data-idvar");
					console.log("elenco_movimenti:"+idmag);
					$.ajax({
						url     : "pag_adm_magazzino.asp?elenco_movimenti="+idmag,
						type    : "post",
						cache: false,
						//dataType: 'json',
						//data	: dati,
						success: function(data){
							$("#elenco_movimenti").html(data);
						}
						,error: function(xhr, textStatus, error){
							toastr.error('Errore nel caricamento elenco movimenti','',{timeOut: 0});
							
							var txt="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
							invia_errore("Errore nel caricamento elenco movimenti",stringa_dati+txt );
						}
					});	
				}
				else{

				var win = window.open("pag_adm_movimenti.asp?idpro="+idpro+idvar, '_blank');
				}
				break

			}
		},
		// Implement the beforeOpen callback to dynamically change the entries
		beforeOpen: function(event, ui) {
			var $menu = ui.menu,
				$target = ui.target;	
				var escludi=$target.attr("data-menu-escludi");
				console.log("escludi:"+escludi);
				
				switch (escludi){
					case "g":
					$(document).contextmenu_idpro("showEntry", "giacenze", false);
				}
				
				
						
		//				.contextmenu("replaceMenu", [{title: "aaa"}, {title: "bbb"}])
		//				.contextmenu("replaceMenu", "#options2")
		//				.contextmenu("setEntry", "cut", {title: "Cuty", uiIcon: "ui-icon-heart", disabled: true})
		// Optionally return false, to prevent opening the menu now
		}
	});	
	
	    $(document).on("click",".dipendente", function(e) {
        var posizione_errore = ".dipendente.click";
        e.preventDefault();
        var iddip = $(this).attr("id");
        iddip=iddip.replace('dipendente', '');
        console.log("dipendente:" + iddip);
        txtiduser="";
        if (typeof iduser !== 'undefined'){
	        txtiduser="&iduser="+iduser;
        }
        $("#dialog").html("<center>Attendi...</center>");
        $.ajax({
            url: "pag_adm_user.asp?tab=16&dialog=si&iddip=" + iddip+txtiduser,
            type: "post",
            cache: false,
            //dataType: 'json',
            //data	: dati,
            success: function(data) {
                $("#dialog").html(data);
                $("#dialog").dialog({
                    autoOpen: true,
                    modal: true,
                    resizable: true,
                    width: "auto",
                    height: "auto",
                    maxHeight: $(window).height() * 0.9,
                    title: "Dati dipendente",
                    position: {
                        my: "center",
                        at: "center",
                        of: window
                    }
                });
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in pulsante_modifica_intestazione()", txt);
            }
        });
    });


})
function aggiorna_tabella_incassi( idDocumento,idfat){
	posizione_errore="aggiorna_tabella_incassi";
	console.log("aggiorna_tabella_incassi(idord:"+ idDocumento+", idfat:"+idfat);
	//return true;
	var dati = {
	    idord:    	 idDocumento,
		idfat:	idfat,
	    totale:		"null"
	}
	$.ajax({
			url     : "ajax_function.asp?oper=aggiorna_incassi",
			type    : "post",
			cache: false,
			dataType: 'html',
			data	: dati,
			success: function(data){
				//$(".div_tabella_incassi").html(data);
				
				 $(".div_tabella_incassi").each(function () {
					var self = $(this);
					self.html(data);
				  });
			}
			,error: function(xhr, textStatus, error){
				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
				toastr.error("Errore nell'aggiornamento dei dati",'',{timeOut: 0});
				
				var txt="Errore in <%=questofile%>: aggiorna_tabella_incassi("+ idDocumento+","+idfat+")";
				txt+="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
				invia_errore("Errore in aggiorna_tabella_incassi()",txt );				
			}
		});	
	return true;
}
function aggiorna_tabella_scadenze(idfat, idDocumento){
	posizione_errore="aggiorna_tabella_scadenze";
	console.log("aggiorna_tabella_scadenze(idfat:"+idfat);
	//return true;
	var dati = {
		idfat:	idfat,
		idord:  idDocumento,
	    totale:		"null"
	}
	$.ajax({
			url     : "ajax_function.asp?oper=aggiorna_scadenze",
			type    : "post",
			cache: false,
			dataType: 'html',
			data	: dati,
			success: function(data){
				//$(".div_tabella_incassi").html(data);
				
				 $(".div_tabella_scadenze").each(function () {
					var self = $(this);
					self.html(data);
				  });
			}
			,error: function(xhr, textStatus, error){
				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
				toastr.error("Errore nell'aggiornamento dei dati",'',{timeOut: 0});
				
				var txt="Errore in <%=questofile%>: aggiorna_tabella_incassi("+ idDocumento+","+idfat+")";
				txt+="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
				invia_errore("Errore in aggiorna_tabella_incassi()",txt );				
			}
		});	
	return true;
}
function aggiorna_tabella_scadenze_for(idfat){
	posizione_errore="aggiorna_tabella_scadenze_for";
	console.log("aggiorna_tabella_scadenze_for(idfat:"+idfat);
	//return true;
	var dati = {
		idfat:	idfat
	}
	$.ajax({
			url     : "ajax_function.asp?oper=aggiorna_scadenze_for",
			type    : "post",
			cache: false,
			dataType: 'html',
			data	: dati,
			success: function(data){
				
				 $("#div_tabella_scadenze_for").each(function () {
					var self = $(this);
					self.html(data);
				  });
			}
			,error: function(xhr, textStatus, error){
				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
				toastr.error("Errore nell'aggiornamento dei dati",'',{timeOut: 0});
				
				var txt="Errore in <%=questofile%>: aggiorna_tabella_incassi("+ idDocumento+","+idfat+")";
				txt+="xhr.statusText : " + xhr.statusText+"<br>xhr.responseText : " + xhr.responseText+"<br>textStatus : " + textStatus+"<br>error : " + error+"<br>Useragent : "+navigator.userAgent;
				invia_errore("Errore in aggiorna_tabella_incassi()",txt );				
			}
		});	
	return true;
}


function DialogYesNo(testo,func) {
    $("#dialog").html(testo);
	console.log("qui");
    // Define the Dialog and its properties.
    
    
    $("#dialog").dialog({
	    autoOpen: true,
        resizable: false,
        modal: true,
        title: "Richiesta conferma",
        height: "auto",
        width: "auto",
        buttons: {
            
                "No": function () {
                $(this).dialog('close');
                //callback(false);
            },
            "Si": function () {
                $(this).dialog('close');
                func();
            }
        }
    });
}

function elimina_ordine(){
    	$("#operazione").val('elimina_ordine');
	    document.getElementById('form1').submit();
}

function annulla_ordine(){
    	$("#operazione").val('annulla_ordine');
	    document.getElementById('form1').submit();	
}

function f_elimina_fattura(){
    	//$("#idfat").val('elimina_fattura');
	    //$("#form-fatt").submit();
	    location.href='pag_adm_fatture.asp?elimina_fattura='+idfat;
	    
	    
}
function f_elimina_ddt(){
    	//$("#elimina_ddt").val('elimina_ddt');
	    //$("#form-ddt").submit();
	    location.href='pag_adm_ddt.asp?elimina_ddt='+ idDocumento;
}
function f_elimina_notacredito(){
    	//$("#elimina_ddt").val('elimina_ddt');
	    //$("#form-ddt").submit();
	    location.href='pag_adm_note_credito.asp?elimina_nota='+ idDocumento;
}






var menu;






function puls_giu() {
    $(".puls_giu").button({
        text: false,
        icons: {
            primary: "ui-icon-triangle-1-s"
        }
    }).click(function() {
        menu = $(this).parent().next().show().position({
            my: "right top",
            at: "right bottom",
            of: this
        });
        $(".puls_giu").click(function() {
            menu.toggle();
        });
        return false;
    }).parent().buttonset().next().hide().menu();
    $(".puls_giu").show();

}
