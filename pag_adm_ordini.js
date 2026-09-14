$(function() {

    $.ajaxSetup({
        cache: false
    });


    if(typeof inattivo !=='undefined'){
    //Gestione Idle utente
	    var awayCallback = function() {
	        if (idadmin == 1) {
	            //Solo per me
	            toastr.options = {
	                "positionClass": "toast-bottom-right",
	                "timeOut": "0",
	                "extendedTimeOut": "0",
	                "closeButton": true
	            };
	            toastr.success("Inattivo", "Messaggio debug");
	        }
	        inattivo = true;
	    };
	    var idle = new Idle({
	        onAway: awayCallback,
	        awayTimeout: t_awayTimeout //away with default value of the textbox
	    });
	    //Gestione Prenotazione

	    if (prenotazione.indexOf("ok") >= 0) {
	        if (idadmin == 1) {
	            //Risposta refresh solo per me
	            toastr.success(prenotazione, "Messaggio debug",{timeOut: 5000});
	        }
	        //Gestione refresh prenotazione
	        $.ajaxSetup({
	            cache: false
	        });
	        var refreshId = setInterval(function() {


	            $.ajax({
	                url: "record_locking_ajax.asp?tabella="+tabella(documento)+"&id_tabella=" + idDocumento + "&inattivo=" + inattivo,
	                type: "get",
	                timeout: 5000,
	                cache: false,

	                //dataType: 'json',
	                //data	: dati,
	                success: function(result) {
	                    if (result != "") {
	                        if (result.indexOf("ko3") >= 0) {
	                            location.href = questofile + "?avviso=1";
	                        }
	                        if (idadmin == 1) {
	                            //Risposta refresh solo per me
	                            toastr.success(result, "Messaggio debug");
	                        }
	                    } else {
	                        toastr.error('Errore in aggiornamento prenotazione ordine','',{timeOut: 0});
	                    }


	                },
	                error: function(xhr, textStatus, error) {
	                    //if(textStatus==="timeout") {
	                    clearInterval(refreshId);
	                    toastr.options = {
	                        "positionClass": "toast-bottom-right",
	                        "timeOut": "0",
	                        "extendedTimeOut": "0",
	                        "closeButton": true
	                    };
	                    toastr.error('Errore in aggiornamento prenotazione ordine','',{timeOut: 0});

	                    //}

	                }
	            });
	        }, rl_t_refresh);
	    } else {
	        toastr.options = {
	            "positionClass": "toast-bottom-right",
	            "timeOut": "0",
	            "extendedTimeOut": "0",
	            "closeButton": true
	        };
	        toastr.warning(prenotazione);
	    }

}



    $(document).contextmenu({
        delegate: ".left_menu",
        preventSelect: true,
        apertura: "click",
        taphold: true,
        menu: [{
            title: "Modifica spettanze",
            cmd: "edit"
        }, {
            title: "Nota su articolo",
            cmd: "nota",
            uiIcon: "ui-icon-note"
        },{
            title: "Aggiungi riga sopra",
            cmd: "sopra",
            uiIcon: "ui-icon-arrowstop-1-n"
        }, {
            title: "Aggiungi riga sotto",
            cmd: "sotto",
            uiIcon: "ui-icon-arrowstop-1-s"
        }, {
            title: "----"
        }, {
            title: "Elimina riga",
            cmd: "elimina",
            uiIcon: "ui-icon-trash"
        }, {
            title: "----"
        }, {
            title: "Seleziona",
            cmd: "seleziona",
            uiIcon: "ui-icon-check"
        }, {
            title: "Deseleziona",
            cmd: "desel"
        }, {
            title: "----"
        }, {
            title: "Segna Ordinato",
            cmd: "ordinato"
        }],
        select: function(event, ui) {
            var $target = ui.target;
            var iddett = $target.attr("id");
            var ordine = $target.attr("data-ordine");

            switch (ui.cmd) {
                case "nota":
                	console.log("apri");
                    apri_dialog_nota(iddett);
                    break;


                case "sopra":
                    ordine -= 2;
                    apri_dialog_aggiungi(ordine);
                    break;
                case "sotto":
                    apri_dialog_aggiungi(ordine);
                    break;
                case "elimina":
                    $.ajaxSetup({
                        cache: false
                    });
                    $.get("pag_adm_ordini_ajax.asp?documento=" + documento + "&oper=tbody&idord=" + idDocumento + "&elimina_riga_ordine=elimina&nord=" + nDocumento + "&iddett=" + iddett, function(result) {
                        $("#tab_articoli_body").remove();
                        $("#tab_articoli_totali").remove();
                        $('#tab_articoli_testa').after(result);
                        toastr.options = {
                            "positionClass": "toast-bottom-right"
                        };
                        toastr.success('Riga eliminata');
                    });
                    break;
                case "seleziona":
                    $("#sel_" + iddett).html('<span class="ui-icon ui-icon-check ui-corner-all handle bg-gray" ></span>');
                    $("#selcs_" + iddett).val('O');
                    break;
                case "desel":
                    $("#sel_" + iddett).html('');
                    $("#selcs_" + iddett).val('');
                    break;
                case "ordinato":
                    $("#sel_" + iddett).text('Ordinato');
                    $("#selcs_" + iddett).val('1');
                    break;
                case "edit":
                    apri_dialog_edit(iddett);
                    break;


            }
        },
        // Implement the beforeOpen callback to dynamically change the entries
        beforeOpen: function(event, ui) {
            var $menu = ui.menu,
                $target = ui.target;
            var iddett = $target.attr("id");
            console.log("iddett:"+iddett);

            if (spettanze == 0) {
                $(document).contextmenu("showEntry", "edit", false);
            }

            var visibile = true;
            if ($("#sel_" + iddett).text() == "Ordinato") {
                visibile = false;
            }
            $(document).contextmenu("showEntry", "ordinato", visibile);
            $(document).contextmenu("showEntry", "nota", documento!="fattura");
            //$(document).contextmenu("showEntry", "nota", tabella=="ordini"?true:false);
            //				.contextmenu("replaceMenu", [{title: "aaa"}, {title: "bbb"}])
            //				.contextmenu("replaceMenu", "#options2")
            //				.contextmenu("setEntry", "cut", {title: "Cuty", uiIcon: "ui-icon-heart", disabled: true})
            // Optionally return false, to prevent opening the menu now
        }
    });
    $("#dialog").bind('dialogclose', function(event) {
        if (aggiorna_ordine) {
            $("#Aggiorna_ordine").trigger("click");
        }
    });
    var iframe = $('<iframe frameborder="0" marginwidth="0" marginheight="0" allowfullscreen></iframe>');
    var dialog = $("#dialog").append(iframe).dialog({
        autoOpen: false,
        modal: true,
        resizable: false,
        width: "auto",
        height: "auto",
        close: function() {
            iframe.attr("src", "");
        }
    });
    $("#puls_aggiungi_articolo").on("click", function(e) {
        aggiorna_ordine = false;
        e.preventDefault();
        apri_dialog_aggiungi(-1);


    });





    function apri_dialog_aggiungi(ordine) {
        posizione_errore = "apri_dialog_aggiungi";
        $("#dialog").html("Attendi...");
        $.ajax({
            url: "pag_adm_dialog_ordini_aggiungi.asp",
            type: "post",
            cache: false,
            //dataType: 'json',
            data	: {
	            documento: documento,
	            idDocumento: idDocumento,
	            nDocumento: nDocumento,
	            ordine: ordine
            },
            success: function(data) {
                $("#dialog").html(data);
                $("#dialog").dialog({
		            autoOpen: true,
		            modal: true,
		            resizable: false,
		            width: 600,
		            height: "auto",
		            title: "Aggiungi articoli "
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
        return true;
    }

      function apri_dialog_nota(iddett) {
		console.log("apri_dialog_nota");
		var testo=$("#notad"+iddett).html();
		if(testo===undefined){
			testo='';
		}
		else{
			testo = testo.replace(/<br>/gi, "\n");
		}

		console.log("testo:"+testo);
        posizione_errore = "apri_dialog_nota";
        $("#dialog").html('<textarea id="nota'+iddett+'" style="width:450px; height:300px;">'+testo+'</textarea>');
                    var buttons = {
                        "Salva": function() {
                            salvaNota(iddett);
                        }
                      };
            $("#dialog").dialog(

                'option',
                'buttons', buttons
            );
            $("#dialog").dialog({
                autoOpen: true,
	            modal: true,
	            resizable: false,
                width: "auto",
                maxHeight: $(window).height() * 0.9,
                title: "Aggiungi nota"

            });
        return true;
    }
	function salvaNota(iddett){
		dati={
			idord: idDocumento,
			nord: nDocumento,
			iddett: iddett,
			tabella: tabella(documento),
			nota: $("#nota"+iddett).val(),
			oper:"salva_nota"
		}
        $.ajax({
            url: "ajax_function.asp",
            type: "post",
            cache: false,
            dataType: 'json',
            data	: dati,
            success: function(data) {
				aggiorna_tabella("");
                $("#dialog").html("");
                $("#dialog").dialog({
                    buttons: {} //Rimuove pulsanti di precedenti chiamate
                });
                $("#dialog").dialog('close');

            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});
                console.log("xhr.statusText:" + xhr.statusText);
                console.log("xhr.responseText:" + xhr.responseText);
                console.log("textStatus:" + textStatus);
                console.log("error:" + error);
            }
        });


	}


    function apri_dialog_edit(iddett) {
        posizione_errore = "apri_dialog_edit";
        $("#dialog").html("Attendi...");
        $.ajax({
            url: "pag_adm_dialog_ordini_edit.asp?idord=" + idDocumento + "&nord=" + nDocumento + "&iddett=" + iddett + "&tabella=" + tabella(documento),
            type: "post",
            cache: false,
            //dataType: 'json',
            //data	: dati,
            success: function(data) {
                $("#dialog").html(data);
                $("#dialog").dialog({
                    autoOpen: true,
                    modal: true,
                    resizable: false,
                    width: 600,
                    height: "auto",
                    title: "Modifica spettanze ",
                    buttons: {} //Rimuove pulsanti di precedenti chiamate
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


    }

	//Pulsanti
    $("#aggiungi_seriali").click(function(e) {
        e.preventDefault();
        posizione_errore = "aggiungi_seriali";
        $("#dialog").html("Attendi...");
        $.ajax({
            url: "pag_adm_dialog_seriali.asp?idord=" + idDocumento + "&nord=" + nDocumento,
            type: "post",
            cache: false,
            //dataType: 'json',
            //data	: dati,
            success: function(data) {
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
            title: "Aggiungi numeri di serie "
        });
    });
    $("#dialogrollback").click(function(e) {
        e.preventDefault();
        dialogrollback = "aggiungi_seriali";
        $("#dialog").html("Attendi...");
        $.ajax({
            url: "pag_adm_dialog_rollback.asp?idord=" + idDocumento ,
            type: "post",
            cache: false,
            //dataType: 'json',
            //data	: dati,
            success: function(data) {
                $("#dialog").html(data);
		        $("#dialog").dialog({
		            autoOpen: true,
		            modal: true,
		            resizable: false,
		            width: 600,
		            height: "auto",
		            title: "Elenco backup spettanze "
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
    });



    $("#mieidettagli").click(function(e) {
        e.preventDefault();
        $(".mieidettagli").toggle();
    });

    $("#prova").click(function(e) {
        console.log("prova");
        e.preventDefault();
        $("#dialog").html("Attendi...");
        $.ajax({
            cache: false,
            url: "pag_adm_dialog_mail.asp?idord=" + idDocumento + "&nord=" + nDocumento,
            type: "post",
            cache: false,
            //dataType: 'json',
            //data	: dati,
            success: function(data) {
                $("#dialog").html(data);
                $("#dialog").dialog({
                    autoOpen: true,
                    modal: true,
                    resizable: false,
                    width: 600,
                    height: "auto",
                    title: "Invio email al cliente"
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
    });



	$(document).on( "click", ".pulsante_modifica_intestazione", function(e){
//	$(".pulsante_modifica_intestazione").on("click", function(e){

  //  $("#pulsante_modifica_intestazione").click(function(e) {
	    var idDocumento= $(this).attr("data-id");
        posizione_errore = "pulsante_modifica_intestazione";
        var documento=$(this).attr("data-documento");
	    console.log("idord:"+ idDocumento+" documento:"+documento);
	    tipo_documento=documento;
	    titolo_finestra=documento;
        $("#dialog").html("<center>Attendi...</center>");


        $.ajax({
            url: "pag_adm_dialog_ordine.asp?idord=" + idDocumento + "&documento=" + documento + "&cancellabile=<%=cancellabile%>",
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
                    height: "auto",
                    title: "Modifica dati "+titolo_finestra
                });
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});

                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in pulsante_modifica_intestazione()", txt);
            }
        });
    });
    $("#pulsante_qr").click(function(e) {
        posizione_errore = "pulsante_qr";
        e.preventDefault();
        $("#dialog").html('<div id="qr_ordine"></div>');
        jQuery('#qr_ordine').qrcode({
            text: "ordine:" + idDocumento,
            width: 128,
            height: 128
        });
        $("#dialog").dialog({
            autoOpen: true,
            modal: true,
            resizable: true,
            width: "auto",
            height: "auto",
            title: "Codice QR"
        });
    });

    $("#pulsante_crea_ddt_fattura").click(function(e) {

	    if (typeof ddt_o_fattura !== 'undefined') {
		    cosa=ddt_o_fattura;
		}
		else{
			cosa='';
		}
		switch (cosa){
			case 'D':
				titolo_finestra="ddt";
				break;
			case 'F':
				titolo_finestra="fattura";
				break;
			default:
				titolo_finestra="ddt o fattura";

		}
        posizione_errore = "pulsante_crea_ddt_fattura";
        $("#dialog").html("Attendi...");
        $.ajax({
            url: "pag_adm_dialog_creafattura.asp?idord=" + idDocumento + "&nord=" + nDocumento,
            type: "post",
            cache: false,
            //dataType: 'json',
            data	: {
	            cosa: cosa,
            	ddt_o_fattura: ddt_o_fattura
            },
            success: function(data) {
                $("#dialog").html(data);
		        $("#dialog").dialog({
		            autoOpen: true,
		            modal: true,
		            resizable: true,
		            width: 600,
		            height: 'auto',
		            title: "Crea "+titolo_finestra
		        });
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});

                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in pulsante_crea_ddt_fattura()", txt);


            }
        });
    });
    $("#salva_nota").click(function(e) {
        posizione_errore = "salva_nota";
        var note_gestore = tinyMCE.get('note_gestore').getContent({
            format: 'html'
        });
        var oper = $(this).attr("value") + "_nota";

        var dati = {
            nota: note_gestore
        }


        console.log(oper);
        $.ajax({
            url: "pag_adm_ordini_ajax.asp?tabella=" + tabella(documento) + "&oper=" + oper + "&idord=" + idDocumento,
            type: "post",
            cache: false,
            //dataType: 'json',
            data: dati,
            success: function(data) {
                $("#cella_note_gestore").html(data);
                $("#note").show();
                //showhide_tabella_note();
                toastr.success('Nota salvata.');
                tinyMCE.get('note_gestore').setContent("");
                console.log(data);
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore in aggiungi o modifica nota','',{timeOut: 0});


            }
        });


    });

    $("#crea_sostituzione").click(function(e) {
        var stringa_selezionati = "";
        var _ordina = false;
        $('.selezione').each(function() {
            if ($(this).val() == "O") {
                if (stringa_selezionati != "") {
                    stringa_selezionati += ",";
                }
                stringa_selezionati = stringa_selezionati + $(this).attr("id").replace("selcs_", "");
                _ordina = true;
            }
        });
		location.href="pag_adm_user.asp?crea_sostituzione2=si&idord="+ idDocumento;

		console.log("Creo sostituzione per ordine:"+ idDocumento+" stringa_selezionati:"+stringa_selezionati);




    });




    $("#allega_files").on("click", function(e) {
        aggiorna_ordine = true;
        e.preventDefault();

        var src = "pag_upload_allegati_new.asp?tipo_allegato=" + tipo_allegato + "&id_tipo_allegato=" + idDocumento;
        var title = "Aggiungi files a ordine " + nDocumento;
        var width = 600;
        var height = 450;
        iframe.attr({
            width: +width,
            height: +height,
            src: src
        });
        dialog.dialog("option", "title", title).dialog("open");
    });
    $("#cambia_intestazione").on("click", function(e) {
        e.preventDefault();

        $("#dialog").html("<center>Attendi...</center>");
        $.ajax({
            url     : "pag_admin_cambia_intestazione.asp?id_documento=" + idDocumento+"&documento="+documento,
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
            height: 200,
            title: "Cambia intestazione documento"
        });
    });
    $("#allega_files2").on("click", function(e) {
        e.preventDefault();
        $("#dialog").html('Attendi');
        $.ajax({
            url: "sub_dropzone.asp",
            //Spostare sotto dati tipo_allegato=" + tipo_allegato + "&id_tipo_allegato=" + idDocumento
            type: "post",
            cache: false,
            //dataType: 'json',
            data: { tipo_allegato: tipo_allegato,id_tipo_allegato: idDocumento},
            success: function(data) {
	            $("#dialog").html(data);
				$("#dialog").dialog({
		            autoOpen: true,
		            modal: true,
		            resizable: true,
		            width: 'auto',
		            height: 'auto',
		            title: "Aggiungi files a ordine " + nDocumento
		        });
            },
            error: function(xhr, textStatus, error) {
                toastr.error('Errore nel caricamento','',{timeOut: 0});


            }
        });


    });




    $(".oper_articoli").on("click", function(e) {
        e.preventDefault();
        var operazione = $(this).attr("id");
        console.log("Premuto pulsante " + operazione);
        var titledialog;
        var bodydialog;
        switch (operazione) {
             case "sostituzione":
                titledialog = "Sostituzione articoli";
                bodydialog = "Scegli gli articoli da ordinare ai fornitori";
                //Create the buttons object
                var buttons = {
                    "Articoli selezionati": function() {
                        elenco_articoli_ordina('selezionati', stringa_selezionati, operazione);
                    },
                    "Vuota": function() {
                        elenco_articoli_ordina('mancanti', "", operazione);
                    }
                };
                break;





            case "ordina_fornitore":
                titledialog = "Ordina a fornitore";
                bodydialog = "Scegli gli articoli da ordinare ai fornitori";
                //Create the buttons object
                var buttons = {
                    "Articoli selezionati": function() {
                        elenco_articoli_ordina('selezionati', stringa_selezionati, operazione);
                    },
                    "Articoli mancanti": function() {
                        elenco_articoli_ordina('mancanti', stringa_selezionati, operazione);
                    },
                    "Tutti gli articoli": function() {
                        elenco_articoli_ordina('tutti', stringa_selezionati, operazione);
                    }
                };
                break;

            case "sposta_articoli":
                titledialog = "Sposta o copia gli articoli";
                bodydialog = "Scegli gli articoli da spostare o copiare su nuovo " + cosa;

                //Create the buttons object
                var buttons = {
                    "Articoli selezionati": function() {
                        elenco_articoli_ordina('selezionati', stringa_selezionati, operazione);
                    },
                    "Articoli mancanti": function() {
                        elenco_articoli_ordina('mancanti', stringa_selezionati, operazione);
                    },
                    "Tutti gli articoli": function() {
                        elenco_articoli_ordina('tutti', stringa_selezionati, operazione);
                    }
                };
                break;

            case "seleziona_articoli":
                titledialog = "Seleziona gli articoli";

                elenco_articoli_ordina('tutti', '', 'seleziona_articoli');




                bodydialog = "";


                var buttons = { };
                break;

            case "duplica":
                titledialog = "Duplica " + cosa + " su nuovo cliente";
                bodydialog = '<form action="' + window.location.href + '" method="post" id="form-duplica"><input type="hidden" name="iduser_duplica_hidden" id="iduser_duplica_hidden" value="" style="width: 300px;"/><input type="hidden" name="idord" value="' + idDocumento + '"><input type="hidden" name="operazione" value="duplica"></form>'
                bodydialog += "<script>$(document).ready(function () {$('#iduser_duplica_hidden').select2({placeholder: 'Seleziona cliente',minimumInputLength: 1,width: 'resolve',ajax: {quietMillis: 150,url: 'ajax_function.asp?select2=utenti_sel2',dataType: 'json',data: function (term, page) {return {term: term};},results: function (data) {return {results: data};}}});});</script>"
                var buttons = {
                    "Duplica": function() {
                        var newiduser = $("#iduser_duplica_hidden").val();
                        console.log("newiduser:" + newiduser);
                        $.ajax({
                            url: "pag_adm_dialog_ordine.asp",
                            type: "post",
                            cache: false,
                            //dataType: 'json',
                            data: {
                                duplicaidord: idDocumento,
                                cosa: cosa,
								documento: documento,
                                iduser: newiduser,
                            },
                            success: function(data) {
                                $("#dialog").html(data);
                                $("#dialog").dialog({
                                    width: 600,
                                    height: "auto",
                                    title: "Duplica " + cosa + " " + nDocumento,
                                    buttons: {}
                                });
                            },
                            error: function(xhr, textStatus, error) {
                                toastr.error('Errore nel caricamento della pagina','',{timeOut: 0});

                                var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                                invia_errore("Errore in pulsante_crea_ddt_fattura()", txt);
                            }
                        });

                    }
                };


                break;
        }



        $("#dialog").html(bodydialog);


        //Conto selezionati per ordine a fornitore
        var stringa_selezionati = "";
        var _ordina = false;
        $('.selezione').each(function() {
            if ($(this).val() == "O") {
                if (stringa_selezionati != "") {
                    stringa_selezionati += ",";
                }
                stringa_selezionati = stringa_selezionati + $(this).attr("id").replace("selcs_", "");
                _ordina = true;
            }
        });
        //Aggiungo il pulsante
        if (_ordina) {
            console.log("stringa_selezionati:" + stringa_selezionati);
        } else {
            delete buttons["Articoli selezionati"];

        }
        $("#dialog").dialog({
            autoOpen: true,
            modal: true,
            resizable: false,
            width: 600,
            height: "auto",
            title: titledialog,
            buttons: buttons
        });
    });

    function elenco_articoli_ordina(filtro, selezionati, operazione) {
        var dati = {
            idord: idDocumento,
            filtro: filtro,
            selezionati: selezionati,
            operazione: operazione,
            tabella: tabella(documento)
        };
        $.ajax({
            url: "pag_adm_dialog_oper_artic.asp",
            type: "post",
            cache: false,
            dataType: 'html',
            data: dati,
            success: function(data) {
                $("#dialog").html(data);
                switch (operazione) {
                    case "ordina_fornitore":
                        var buttons = {
                            "Ordina": function() {
                                invia_elenco_articoli();
                            }
                        };
                        break;
                    case "sposta_articoli":
                        var buttons = {
                            "Sposta su nuovo ordine": function() {
                                invia_elenco_articoli();
                            },
                            "Copia su nuovo ordine": function() {
                                $("#operazione_dialog").val("copia_articoli");
                                console.log("operazione_dialog" + $("#operazione_dialog").val());
                                invia_elenco_articoli();
                            }
                        };
                        break;

                    case "impegna_articoli":
                        var buttons = {
                            "impegna": function() {
                                invia_elenco_articoli();
                            }
                        };
                        break;
                     case "seleziona_articoli":
                        var buttons = {};
                        break;


                }
                $("#dialog").dialog(

                    'option',
                    'buttons', buttons
                );
                $("#dialog").dialog({
                    width: "auto",
                    maxHeight: $(window).height() * 0.9
                });



            },
            error: function(xhr, textStatus, error) {
                toastr.error("Errore nell'aggiornamento dei dati",'',{timeOut: 0});
                var txt = "Errore in <%=questofile%>: ordina_fornitore(" + idDocumento + ")";
                txt += "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in elenco_articoli_ordina()", txt);
            }
        });
    }

    function invia_elenco_articoli() {

        var dati = $("#form_oper_artic").serialize() + "&tabella=" + tabella(documento);
        $.ajax({
            url: "pag_adm_dialog_oper_artic.asp",
            type: "post",
            cache: false,
            dataType: 'json',
            data: dati,
            success: function(data) {


                $.ajaxSetup({
                    cache: false
                });
                $.get("pag_adm_ordini_ajax.asp?tabella=" + tabella(documento) + "&oper=tbody&idord=" + idDocumento + "&nord=" + nDocumento, function(result) {
                    $("#tab_articoli_body").remove();
                    $("#tab_articoli_totali").remove();
                    $('#tab_articoli_testa').after(result);
                });


                $("#dialog").html(data.message);
                $("#dialog").dialog('option', 'buttons', {
                    "Chiudi": function() {
                        $("#dialog").dialog('close');
                    }
                });
            },
            error: function(xhr, textStatus, error) {
                toastr.error("Errore nell'aggiornamento dei dati",'',{timeOut: 0});
                var txt = "Errore in <%=questofile%>: ordina_fornitore(" + idDocumento + ")";
                txt += "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
                invia_errore("Errore in invia_elenco_articoli()", txt);
            }
        });




    }
    $('#tab_articoli_body').sortable({
       placeholder: "ui-state-highlight",
        forcePlaceholderSize: true,
	    forceHelperSize: true,
        helper: fixHelper,
        handle: '.handle',

        start: function(e, ui ){
		     ui.placeholder.height(ui.helper.outerHeight());
 		},
        update: function(event, ui) {
            var order = $(this).sortable('serialize');
            $.get("pag_adm_ordini_ajax.asp?tabella=" + tabella(documento) + "&oper=riordina&idord=" + idDocumento + "&nord=" + nDocumento + "&" + order, function(result) {
                if (result != "") {
                    toastr.options = {
                        "positionClass": "toast-bottom-right"
                    };
                    toastr.success(result);
                } else {
                    toastr.error('Errore in ordinamento elenco','',{timeOut: 0});
                }
            });
        }
    });




}); //fine $(function () {

//Aggiungiarticoli, passare in dati
function aggiorna_tabella(dati) {
    posizione_errore = "aggiorna_tabella";
    console.log(dati);
    //return true;
    $.ajax({
        url: "pag_adm_ordini_ajax.asp?documento=" + documento + "&oper=tbody&idord=" + idDocumento + "&nord=" + nDocumento,
        type: "post",
        cache: false,
        //dataType: 'json',
        data: dati,


        success: function(data) {
            //alert(data.Message);

            $("#tab_articoli_body").remove();
            $("#tab_articoli_totali").remove();
            $('#tab_articoli_testa').after(data);
            aggiorna_ordine = false;
        },
        error: function(xhr, textStatus, error) {
            toastr.error("Errore nell'aggiornamento dei dati",'',{timeOut: 0});


            var txt = "xhr.statusText : " + xhr.statusText + "<br>xhr.responseText : " + xhr.responseText + "<br>textStatus : " + textStatus + "<br>error : " + error + "<br>Useragent : " + navigator.userAgent;
            invia_errore("Errore in aggiorna_tabella()", dati + txt);

        }
    });
    return true;
}


function prendi() {
    posizione_errore = "prendi";
    $.get("record_locking_ajax.asp?tabella=ordine&id_tabella=" + idDocumento + "&prendi=true", function(result) {
        if (result != "") {
            toastr.options = {
                "positionClass": "toast-bottom-right",
                "timeOut": "0",
                "extendedTimeOut": "0",
                "closeButton": true
            };
            toastr.success(result);
        } else {
            toastr.error('Errore in aggiornamento prenotazione ordine','',{timeOut: 0});
        }
    });
}

// Return a helper with preserved width of cells
var fixHelper = function(e, ui) {
    ui.children().each(function() {
        $(this).width($(this).width());
    });
    return ui;
};

function isNumber(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
}

function isOneSel() {
    var _isonesel = false;
    $('.selezione').each(function() {
        if ($(this).val() != "") {
            _isonesel = true;
        }
    });
    return _isonesel;
}

function isOneChecked() {
    return ($(".selezione:checked").length > 0);
}

function tabella(documento){
	switch (documento){
		case "fattura":
			return "fatture";
			break;
		case "preventivo":
			return "preventivi";
			break;
		case "ordine_fornitore":
			return "ordini_fornitori";
			break;
		default:
			return "ordini";
	}



}


function doAutogrow() {
    $(".autogrow").autogrow({
        onInitialize: true,
        speed: 50
    });
}


