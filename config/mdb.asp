<%
'FILE MDB per gestionale.larochelle.it
const upl_img_cat="/public/foto/"			'directory immagini articoli
const prg_dir="/public/progetti/"		'directory immagini progetti
const host="ARUBA"
const nomesito="gestionale.larochelle.it"		'nome del sito in formato www
const SMTPServer="smtp.aruba.it"  'server smpt per posta in uscita
const col_1=160
const col_2=600
const col_3=160
const col_23=762
const col_4=924
const cellp=1
const cells=0
const size_chr_pdf=7 'dimesnione carattere elenco articoli pdf
const stato_min_ddt=4	'stato minimo per abilitazione fattura e ddt
const reg_utenti=true	'Abilita la registrazione degli utenti
const simbolo_valuta="&euro; "
const mysql_server="62.149.150.169"
const mysql_uid="Sql594349"
const mysql_pwd="a125ee15"
const mysql_database="Sql594349_2"
const pw_manuali="pw"			'Password pagina manuals.asp
const mod_dip=true	'Abilita tutta la parte dipendenti
const mod_larochelle=true	'Altre modifiche solo LaRochelle



const dimensione_logo_fattura=170

' Confistring
' -1- Chat livezilla
' -2- File mager fmsmall
' -3- File mager KCfinder
const configstring="-1-2-3-"
const edit_lang_en=false
const lang_en=false
function stato_ordine(stato)
	if not isnumeric(stato) then
		stato_ordine=stato
	else
		select case stato
		case 0
		stato_ordine="Nuovo ordine in controllo disponibilit&agrave;"
		case 1
		stato_ordine="In modifica"
		case 2
		stato_ordine="Pronto per lavorazione"
		case 3
		stato_ordine="In lavorazione"
		case 4
		stato_ordine="Anticipati"
		case 5
		stato_ordine="Parziale consegnato"
		case 6
		stato_ordine="Evaso"
		case 7
		stato_ordine="Annullato"
		end select
	end if
end function

function intestazione_ddt_fat(data_documento)
	if DateDiff("d","21/09/2021",data_documento)<0 then 'prima del 22/09/2022
		intestazione_ddt_fat="<b>La Rochelle di Pistono Ilaria & C. snc</b><br>Via Monsignor A. Sangiorgio 59 - 10090 San Giorgio Canavese (TO)<br>P.IVA 06315200011<br>Licenza EX. ART. 28 T.U.L.P.S. - Prot. 0077320/2014 AREA  I TER"
	elseif DateDiff("d","30/05/2023",data_documento)<0 then 'prima del 30/05/2023
		intestazione_ddt_fat="<b>La Rochelle di Pistono Ilaria & C. snc</b><br>Via Paschetto 46 - 10090 San Giorgio Canavese (TO)<br>P.IVA 06315200011<br>Licenza EX. ART. 28 T.U.L.P.S. - Prot. 0077320/2014 AREA  I TER"
     	else
	 'Dopo
		intestazione_ddt_fat="<b>La Rochelle srl</b><br>Via Paschetto 46 - 10090 San Giorgio Canavese (TO)<br>P.IVA 12885200019<br>Licenza EX. ART. 28 T.U.L.P.S. - Prot. 0077320/2014 AREA  I TER"
	end if
end function

function pre_fattura(data_documento)
	'if DateDiff("d","01/01/2016",data_documento)<0 then 'prima del 01/01/2016
	'	pre_fattura="RI-"
	'else 'Dopo
		pre_fattura=""
	'end if
end function
txt_restituzione="<b>Resi</b><br>Si avvisa la clientela che le richieste di riparazioni o sostituzioni del materiale si accettano fino ad un mese dalla consegna. Le sostituzioni, per ragioni fiscali, non possono essere effettuate con articolo differente ma solo con lo stesso codice articolo. Gli articoli dovranno essere integri, muniti dei rispettivi cartellini identificativi del prodotto ed all'interno delle confezioni originali non danneggiate."

sub intestazione_pdf_ddt_fat(byref Doc, ByRef page ,data_documento)
		const file_logo="/public/files/logo_documenti.jpg"
		Set fs = Server.CreateObject("Scripting.filesystemObject")
		if fs.FileExists(Server.MapPath( file_logo)) then
			esiste_logo=true

			'Parte nuova
			dimensione_logo=170
			margine=5
			Set logo = Doc.OpenImage( Server.MapPath( file_logo) )
			'calcolo misure dell'immagine considerando la risoluzione DPI
			larghezza=logo.width*72/logo.resolutionx
			altezza=logo.height*72/logo.resolutiony


			scala=(dimensione_logo)/larghezza

			'calcolo nuova altezza considerando la scala per centrare poi l'immagine
			nuova_altezza=altezza*scala
			nuova_larghezza=larghezza*scala


			param_x =((Page.Width/ 2)-nuova_larghezza)/2

			'Sotto laschio metà del margine
			param_y = Page.Height - nuova_altezza - 10


			param_ScaleX = scala
			param_ScaleY = scala
			parametri="x="&param_x&"; y="&param_y&"; ScaleX="&param_ScaleX&"; ScaleY="&param_ScaleY&";"
			parametri=replace(parametri,",",".")

		else
			param_Y=Page.Height-20

		end if
		Set fs = Nothing

		'response.write parametri
		if esiste_logo then Page.Canvas.DrawImage logo, parametri


		param_Y=param_Y
		param_y=replace(param_y,",",".")
		Page.Canvas.DrawText intestazione_ddt_fat(data_documento), "X=20; Y="&param_y&"; html=true, size=9;" , doc.fonts("arial")


end sub


%>
