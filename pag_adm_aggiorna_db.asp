<!--#include file="adovbs.inc" -->
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_elenchi.asp" -->
<!--#include virtual="/pag_adm_aggiorna_db_sub.asp" -->
<%
Session.LCID=1040
'Response.CharSet = "UTF-8"
'on error resume next

%>
<!--#include virtual="/db_edit_inc.asp" -->
<%

'				case "integer"			: sFieldType = "SMALLINT"
'				case "long"			: sFieldType = "INTEGER"
'				case "boolean"			: sFieldType = "BIT"
'				case "date"				: sFieldType = "DATE"
'				case "currency"			: sFieldType = "CURRENCY"
'				case "text"			: sFieldType = "VARCHAR"
'				case "memo"		: sFieldType = "TEXT"
'				case "ole"	: sFieldType = "LONGVARBINARY"
'				case "guid"				: sFieldType = "GUID"
'				case "byte"	: sFieldType = "TINYINT"
dim rs_or
dim rs
response.write "Campo [impostazioni].[versione]: "
	sql="select * from impostazioni order by id"

	set rs=conn.execute(sql)
	versione=rs("versione")
	rs.close
	if versione="" or isnull(versione) then
		call imposta_versione(1)
		versione=0
	end if
response.write "Versione iniziale"&versione&"<br>"
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `incassi` CHANGE  `importo`  `importo` DECIMAL( 10, 2 ) NOT NULL")
	call imposta_versione(nuova_versione)'141
end if

nuova_versione=142
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `totale`  `totale` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `totale`  `totale` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `totale`  `totale` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `trasporto`  `trasporto` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `trasporto`  `trasporto` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `trasporto`  `trasporto` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `sconto_ordine`  `sconto_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `sconto_ordine`  `sconto_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `sconto_ordine`  `sconto_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `costo_totordine`  `costo_totordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `spese_bancarie_ordine`  `spese_bancarie_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `arrotondamento`  `arrotondamento` DECIMAL( 10, 4 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `totale_merce_ordine`  `totale_merce_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `totale_merce_ordine`  `totale_merce_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `totale_merce_ordine`  `totale_merce_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `imposta_ordine`  `imposta_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `imposta_ordine`  `imposta_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `imposta_ordine`  `imposta_ordine` DECIMAL( 10, 2 ) NOT NULL")
	call imposta_versione(nuova_versione)'142
end if

nuova_versione=143
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `sconto_ordine`  `sconto_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `spese_bancarie`  `spese_bancarie` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `totale_fattura`  `totale_fattura` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `bk_totale_fattura`  `bk_totale_fattura` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `totale_merce`  `totale_merce` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `imposta`  `imposta` DECIMAL( 10, 2 ) NOT NULL")
	call imposta_versione(nuova_versione)'143
end if

nuova_versione=144
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `iva_ordine` INT NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `iva_ordine` INT NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'144
end if

nuova_versione=145
if versione<nuova_versione  then
	conn.execute ("UPDATE  preventivi SET  iva_ordine=22")
	conn.execute ("UPDATE  ordini_fornitori SET  iva_ordine=22")
	call imposta_versione(nuova_versione)'145
end if

nuova_versione=146
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `prezzo`  `prezzo` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `costo`  `costo` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `prezzo_ve_min`  `prezzo_ve_min` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `varianti_a` CHANGE  `prezzo_ac_va`  `prezzo_ac_va` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `varianti_a` CHANGE  `prezzo_ve_va`  `prezzo_ve_va` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `prezzo`  `prezzo` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `sconto_prodotto`  `sconto_prodotto` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `costo_ordine`  `costo_ordine` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `totale_riga`  `totale_riga` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_originale` CHANGE  `prezzo`  `prezzo` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett` CHANGE  `prezzo`  `prezzo` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett` CHANGE  `sconto_prodotto`  `sconto_prodotto` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `totale_riga`  `totale_riga` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett` CHANGE  `prezzo`  `prezzo` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett` CHANGE  `sconto_prodotto`  `sconto_prodotto` DECIMAL( 10, 2 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett` CHANGE  `totale_riga`  `totale_riga` DECIMAL( 10, 2 ) NOT NULL")
	call imposta_versione(nuova_versione)'146
end if

nuova_versione=147
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `calcolato` BOOLEAN NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `calcolato` BOOLEAN NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `calcolato` BOOLEAN NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'147
end if

nuova_versione=148
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `no_magazzino`  `no_magazzino` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `scalato_magazzino`  `scalato_magazzino` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'148
end if

nuova_versione=149
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `Fornitore`  `Fornitore` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `mailing`  `mailing` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `modalita_registrazione`  `modalita_registrazione` SMALLINT UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `non_eliminare`  `non_eliminare` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE `elenco_comuni` ADD FULLTEXT `Comune_FullText` (`Comune`)")
	call imposta_versione(nuova_versione)'149
end if

nuova_versione=150
if versione<nuova_versione  then
	conn.execute ("UPDATE  ordini SET  calcolato=1")
	conn.execute ("UPDATE  preventivi SET  calcolato=1")
	conn.execute ("UPDATE  ordini_fornitori SET  calcolato=1")
	call imposta_versione(nuova_versione)'150
end if

nuova_versione=151
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `attivo`  `attivo` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '1'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `promozione`  `promozione` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `novita`  `novita` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `vetrina`  `vetrina` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `vendita`  `vendita` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '1'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `yatego`  `yatego` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `in_evidenza`  `in_evidenza` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `in_evidenza`  `in_evidenza` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `prezzo_riservato`  `prezzo_riservato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'151
end if
nuova_versione=152
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `PA`  `PA` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `fatture` CHANGE  `fine_mese`  `fine_mese` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'152
end if
nuova_versione=153
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `cap`  `cap` VARCHAR( 10 ) NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'153
end if
nuova_versione=154
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` DROP COLUMN  `scalato_magazzino`")
	call imposta_versione(nuova_versione)'154
end if
nuova_versione=155
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `upc_ean` ( `upc_ean` VARCHAR(13) NOT NULL, `idmag` int(11) DEFAULT NULL, PRIMARY KEY (`upc_ean`)) ENGINE=MyISAM DEFAULT CHARSET=latin1;")
	call imposta_versione(nuova_versione)'155
end if
nuova_versione=156
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `Nord`  `Nord` INT( 11 ) NULL DEFAULT NULL after idord")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `trasporto`  `trasporto` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after totale_merce_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `spese_bancarie_ordine`  `spese_bancarie_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after trasporto")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `imposta_ordine`  `imposta_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after spese_bancarie_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `arrotondamento`  `arrotondamento` DECIMAL( 10, 4 ) NOT NULL DEFAULT  '0' after imposta_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `costo_totordine`  `costo_totordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after totale_merce_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `sconto_ordine`  `sconto_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after arrotondamento")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `totale`  `totale` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after sconto_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `cf`  `cf` VARCHAR( 20 )  NULL DEFAULT '' after piva")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `cig`  `cig` VARCHAR( 12 ) NULL DEFAULT '' after d_cap")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `mepa`  `mepa` VARCHAR( 15 ) NULL DEFAULT '' after cig")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `comunicazioni_precedenti`  `comunicazioni_precedenti` LONGTEXT NULL DEFAULT '' after note_spedizione")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `trattamento_iva_ordine`  `trattamento_iva_ordine` INT( 11 ) NOT NULL DEFAULT  '0' after Tipo_trasporto")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `iva_ordine`  `iva_ordine` INT( 11 ) NOT NULL DEFAULT  '0' after trattamento_iva_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `iva_0`  `iva_0` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after iva_ordine")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `eliminato`  `eliminato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `modificato`  `modificato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `verde`  `verde` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after modificato")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `calcolato`  `calcolato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after verde")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `creato_da_admin`  `creato_da_admin` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after calcolato")	
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `no_magazzino`  `no_magazzino` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after scalato_magazzino")
	call imposta_versione(nuova_versione)'156
end if

nuova_versione=157
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `Nord`  `Nord` INT( 11 ) NULL DEFAULT NULL after idord")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `trasporto`  `trasporto` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after totale_merce_ordine")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `imposta_ordine`  `imposta_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after trasporto")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `sconto_ordine`  `sconto_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after imposta_ordine")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `totale`  `totale` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0' after sconto_ordine")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `cf`  `cf` VARCHAR( 20 )  NULL DEFAULT '' after piva")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `comunicazioni_precedenti`  `comunicazioni_precedenti` LONGTEXT NULL DEFAULT '' after note_spedizione")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `trattamento_iva_ordine`  `trattamento_iva_ordine` INT( 11 ) NOT NULL DEFAULT  '0' after Tipo_trasporto")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `iva_ordine`  `iva_ordine` INT( 11 ) NOT NULL DEFAULT  '0' after trattamento_iva_ordine")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `iva_0`  `iva_0` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after iva_ordine")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `eliminato`  `eliminato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `modificato`  `modificato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `calcolato`  `calcolato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after modificato")
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `creato_da_admin`  `creato_da_admin` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after calcolato")	
	call imposta_versione(nuova_versione)'157
end if

nuova_versione=159
if versione<nuova_versione  then
	conn.execute ("drop table prodotti_ebay")
	conn.execute ("CREATE TABLE IF NOT EXISTS `prodotti_ebay` ( `idebay` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`idebay`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `data_creato` datetime NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `data_modificato` datetime")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `idpro` INT(11) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `site` INT( 3 ) NOT NULL DEFAULT  '101'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `format` INT( 1 ) NOT NULL DEFAULT  '9'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `currency` INT( 2 ) NOT NULL DEFAULT  '7'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `title` VARCHAR( 80 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `SubtitleText` VARCHAR( 80 ) DEFAULT  ''")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `Note` TEXT NULL DEFAULT  ''")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `description` TEXT NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `category_1` INT( 10 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `category_2` INT( 10 ) ")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `store_category` INT( 2 ) ")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `picurl` TEXT")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `quantity` INT( 4 ) NOT NULL DEFAULT  '1'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `duration` INT( 3 ) NOT NULL DEFAULT  '10'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `starting_price` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '1'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `reserve_price` DECIMAL( 10, 2 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `counter` INT( 1 ) NOT NULL DEFAULT  '3'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `payment_instructions` TEXT")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `insurance_option` INT( 4 ) ")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `accept_paypal` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `paypal_email_address` VARCHAR( 80 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `accept_mo_cashiers` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `accept_personal_check` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `accept_visa` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `accept_cod` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `accept_money_xfer` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `location_city_state` VARCHAR( 45 ) NOT NULL DEFAULT  ''")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `location_country` VARCHAR( 2 ) NOT NULL DEFAULT  'IT'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `gallery` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `gallery_plus` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `gallery_featured` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `gallery_url` TEXT NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `bold` TINYINT (1) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `highlight` TINYINT (1) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `featured_plus` TINYINT (1) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `shippingtype` INT( 4 ) NOT NULL DEFAULT  '1'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `shippingpackage` INT( 4 ) ")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `shippingirregular` TINYINT (1) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `weightmajor` INT( 4 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `weightminor` INT( 4 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `packagelength` INT( 4 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `packagewidth` INT( 4 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `packagedepth` INT( 4 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `shipfromzipcode` INT( 12 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `packaginghandligcost` DECIMAL( 10, 2 )")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `autopay` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `apply_multi_item` TINYINT (1) UNSIGNED NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `shippingserviceoptions` TEXT NOT NULL")
	call imposta_versione(nuova_versione)'155
end if

nuova_versione=160
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `scadenze_for` CHANGE  `Pagato`  `Pagato` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'160
end if

nuova_versione=161
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `incassi` CHANGE  `importo`  `importo` DECIMAL( 10, 2 ) NOT NULL")
	call imposta_versione(nuova_versione)'161
end if

nuova_versione=162
if versione<nuova_versione  then
	conn.execute ("UPDATE scadenze SET riba=1 WHERE riba=-1;")
	conn.execute ("ALTER TABLE  `scadenze` CHANGE  `riba`  `riba` TINYINT( 1 ) UNSIGNED NOT NULL")
	call imposta_versione(nuova_versione)'162
end if

nuova_versione=163
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `minimo_preventivo` DECIMAL( 10, 2 ) DEFAULT  '0'")
	call imposta_versione(nuova_versione)'163
end if

nuova_versione=164
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `copia_email_ordini`  VARCHAR( 15 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'164
end if

nuova_versione=165
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` CHANGE  `copia_email_ordini`  `copia_email_ordini` VARCHAR( 100 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'165
end if

nuova_versione=166
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE ordini_dett DROP INDEX iddett")
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `ordinato` TINYINT UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'166
end if

nuova_versione=167
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `magazzino` CHANGE  `db_ven`  `db_ven` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `magazzino` CHANGE  `db_acq`  `db_acq` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'167
end if

nuova_versione=168
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `scalato_magazzino` TINYINT UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'168
end if

nuova_versione=169
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `category_predefinito`  INT( 4 ) NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'169
end if

nuova_versione=170
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `brand` CHANGE  `in_home`  `in_home` TINYINT( 1 ) UNSIGNED NOT NULL")
	call imposta_versione(nuova_versione)'170
end if

nuova_versione=171
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `prezzi_con_iva`  TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'171
end if

nuova_versione=172
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `metodi_pagamento` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `descrizione` VARCHAR( 80 ) NOT NULL ")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `descrizione_checkout` VARCHAR( 255 ) NOT NULL DEFAULT  ''")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `numero_scadenze` INT( 1 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `inizio_scadenze` INT( 1 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `riba` TINYINT (1) UNSIGNED NOT NULL  DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `spese_gestione` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `maggiorazione_gestione` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'172
end if

nuova_versione=173
if versione<nuova_versione  then
	for n=1 to pagamento_fattura(-1,0)
		conn.execute ("INSERT INTO metodi_pagamento (descrizione, numero_scadenze, inizio_scadenze,riba) VALUES ('"&pagamento_fattura(n,0)&"', "&pagamento_fattura(n,2)&", "&pagamento_fattura(n,3)&","&converti_bool(pagamento_fattura(n,4))&")")
	next
	call imposta_versione(nuova_versione)'173
end if

nuova_versione=174
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `in_checkout` TINYINT (1) UNSIGNED NOT NULL  DEFAULT  '0'")
	conn.execute("UPDATE metodi_pagamento SET in_checkout=1 where id=1")
	conn.execute("UPDATE metodi_pagamento SET in_checkout=1 where id=2")
	conn.execute("UPDATE metodi_pagamento SET in_checkout=1 where id=18")
	call imposta_versione(nuova_versione)'174
end if

nuova_versione=175
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_dipendenti` ( `iddip` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`iddip`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `iduser` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `nominativo` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `contatto` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `sesso` VARCHAR( 1 ) NOT NULL DEFAULT 'M'")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `grado` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `nota_dipendente` LONGTEXT CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT ''")
	conn.execute ("CREATE TABLE IF NOT EXISTS `dipendenti_misure` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD `iddip` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_TG` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_fondo` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_manica` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_torace` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_vita` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_spalle` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_lunghezza` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_vita` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_bacino` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `camicia` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `scarpa` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `berretto` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `guanto` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `nota_misura` LONGTEXT CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `data_rilievo` DATETIME NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'175
end if

nuova_versione=176
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_TG` VARCHAR( 10 ) NULL DEFAULT '' AFTER giacca_spalle")
	call imposta_versione(nuova_versione)'176
end if

nuova_versione=177
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` CHANGE  `data_rilievo` `data_rilievo` DATETIME NULL DEFAULT NULL after iddip")
	call imposta_versione(nuova_versione)'177
end if

nuova_versione=178
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_gradi` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `nome_grado` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `iduser` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `giacca_M` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `giacca_F` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `pantaloni_M` VARCHAR( 10 ) NULL DEFAULT '' ")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `pantaloni_F` VARCHAR( 10 ) NULL DEFAULT '' ")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `camicia_M` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `camicia_F` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `scarpa_M` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `scarpa_F` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `berretto_M` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `berretto_F` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `guanto_M` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `guanto_F` VARCHAR( 10 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'178
end if

nuova_versione=179
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `scalato_magazzino`  `scalato_magazzino` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'148
end if

nuova_versione=180
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `numeri_seriali` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `numeri_seriali` ADD  `idpro` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `numeri_seriali` ADD  `idord` INT( 11 )")
	call imposta_versione(nuova_versione)'180
end if

nuova_versione=181
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` ADD  `richiedi_seriale` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti` ADD  `colli` INT( 2 ) ")
	call imposta_versione(nuova_versione)'181
end if

nuova_versione=182
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX idpro ON numeri_seriali (idpro);")
	conn.execute ("CREATE INDEX idord ON numeri_seriali (idord); ")
	call imposta_versione(nuova_versione)'182
end if

nuova_versione=183
if versione<nuova_versione  then
	conn.execute ("UPDATE prodotti SET colli=0")
	conn.execute ("ALTER TABLE preventivi_dett DROP INDEX iddett")
	conn.execute ("ALTER TABLE ordini_fornitori_dett DROP INDEX iddett ")
	conn.execute ("ALTER TABLE fatture DROP INDEX IDFat")
	conn.execute ("ALTER TABLE ddt DROP INDEX idddt")
	conn.execute ("ALTER TABLE carrello DROP INDEX idcar ")
	conn.execute ("ALTER TABLE articoli_consigliati_prodotti DROP INDEX id  ")
	conn.execute ("ALTER TABLE articoli_consigliati DROP INDEX id_pro_consigliati  ")
	conn.execute ("ALTER TABLE settori DROP INDEX idsettore  ")
	
	call imposta_versione(nuova_versione)'183
end if

nuova_versione=184
if versione<nuova_versione  then
	
	conn.execute ("ALTER TABLE  `numeri_seriali` ADD  `seriale` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.Execute("UPDATE numeri_seriali set seriale=id")
	call imposta_versione(nuova_versione)'184
end if
nuova_versione=185
if versione<nuova_versione  then
	
	conn.execute ("ALTER TABLE  `numeri_seriali` CHANGE  `seriale`  `seriale` VARCHAR( 20 ) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'185
end if

nuova_versione=186
if versione<nuova_versione  then
	
	conn.execute ("CREATE UNIQUE INDEX pro_e_var ON magazzino (idpro,idvara,idvarb);")
	conn.execute ("ALTER TABLE magazzino DROP INDEX idvara  ")
	conn.execute ("ALTER TABLE magazzino DROP INDEX idvarb  ")
	call imposta_versione(nuova_versione)'186
end if
nuova_versione=187
if versione<nuova_versione  then
	
	conn.execute ("ALTER TABLE  `magazzino` CHANGE  `idvara`  `idvara` INT( 11 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `magazzino` CHANGE  `idvarb`  `idvarb` INT( 11 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `idord`  `idord` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett` CHANGE  `idord`  `idord` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett` CHANGE  `idord`  `idord` INT( 11 ) NOT NULL")
	conn.execute ("CREATE INDEX pro_e_var ON ordini_dett (idpro,idvara,idvarb);")
	conn.execute ("CREATE INDEX pro_e_var ON ordini_fornitori_dett (idpro,idvara,idvarb);")
	conn.execute ("CREATE INDEX pro_e_var ON preventivi_dett (idpro,idvara,idvarb);")
	call imposta_versione(nuova_versione)'187
end if


nuova_versione=188
if versione<nuova_versione  then
	
	conn.execute ("CREATE TABLE IF NOT EXISTS `prodotti_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD  `iduser` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD  `idpro` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD  `idvara` INT( 11 ) NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD  `idvarb` INT( 11 ) NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD  `tipo_prodotto` INT( 11 ) NOT NULL DEFAULT '0'")
	
	call imposta_versione(nuova_versione)'188
end if
nuova_versione=189
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD  `ordine` INT( 4 ) NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'189
end if
nuova_versione=190
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE `magazzino` DROP `DistintaBase`")
	call imposta_versione(nuova_versione)'190
end if
nuova_versione=191
if versione<nuova_versione  then
	on error resume next
	conn.execute ("ALTER TABLE `magazzino` DROP `variante1`")
	conn.execute ("ALTER TABLE `magazzino` DROP `variante2`")
	call imposta_versione(nuova_versione)'191
end if



nuova_versione=192
if versione<nuova_versione  then
	'conn.execute ("DROP TABLE IF EXISTS prodotti_ebay;")
	
	
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `action` VARCHAR( 30 ) NOT NULL DEFAULT  '' AFTER  `idpro`")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `conditionid` INT( 10 ) NOT NULL ")
	
	
	
	call imposta_versione(nuova_versione)'192
end if

nuova_versione=193
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `impostazioni_ebay` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `Location` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `PayPalEmailAddress` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'193
end if

nuova_versione=194
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti_ebay` DROP COLUMN  `location_city_state`")
	conn.execute ("ALTER TABLE  `prodotti_ebay` DROP COLUMN  `location_country`")
	conn.execute ("ALTER TABLE  `prodotti_ebay` DROP COLUMN  `site`")
	conn.execute ("ALTER TABLE  `prodotti_ebay` DROP COLUMN  `counter`")
	conn.execute ("ALTER TABLE  `prodotti_ebay` DROP COLUMN  `currency`")
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `HitCounter` VARCHAR( 20 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `data_pubblicato` DATETIME NULL DEFAULT NULL after data_modificato")
	
	call imposta_versione(nuova_versione)'194
end if

nuova_versione=195
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `upc_ean` ADD  `idpro` INT( 11 ) DEFAULT NULL")	
	call imposta_versione(nuova_versione)'195
end if

nuova_versione=196
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `DispatchTimeMax` int( 2 ) NOT NULL DEFAULT '1'")
	call imposta_versione(nuova_versione)'196
end if

nuova_versione=197
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `ShippingService_1_Cost` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `ShippingService_1_AdditionalCost` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'197
end if

nuova_versione=198
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `ImmediatePayRequired`  TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'198
end if

nuova_versione=199
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni_ebay` ADD  `store_categories`  LONGTEXT NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'199
end if

nuova_versione=200
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `numeri_seriali` ADD  `iddett` INT( 11 )")
	call imposta_versione(nuova_versione)'198
end if

nuova_versione=201
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `richiedi_seriale`  TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'198
end if

nuova_versione=202
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_dett_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `iddip` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `quantita` INT( 3 ) NOT NULL")
	call imposta_versione(nuova_versione)'202
end if

nuova_versione=203
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `varianti` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `varianti` ADD  `descrizione` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `varianti` ADD  `elenco` TEXT NOT NULL DEFAULT  ''")
	call imposta_versione(nuova_versione)'203
end if

nuova_versione=204
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `itemid` VARCHAR( 30 ) NULL DEFAULT  ''")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `starttime` datetime ")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `endtime` datetime ")
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `status` VARCHAR( 80 )  NULL DEFAULT  ''" )
	conn.execute ("ALTER TABLE  `prodotti_ebay` ADD  `errormessage` TEXT  NULL" )
	call imposta_versione(nuova_versione)'204
end if

nuova_versione=205
if versione<nuova_versione  then
	'conn.execute ("DROP TABLE IF EXISTS prodotti_ebay;")
	conn.execute ("ALTER TABLE  `ordini` ADD  `spettanze` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'205
end if

nuova_versione=206
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX idfat ON scadenze_for (idfat);")
	call imposta_versione(nuova_versione)'206
end if

nuova_versione=207
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `magazzino_movimenti` ADD  `idmag` INT( 11 ) ")
	conn.execute ("ALTER TABLE  `magazzino_movimenti` ADD  `idfor` INT( 11 ) ")
	conn.execute ("CREATE INDEX idmag ON magazzino_movimenti (idmag);")
	call imposta_versione(nuova_versione)'207
end if

nuova_versione=208
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `distinta_base` ADD  `idmag_ven` INT( 11 ) ")
	conn.execute ("ALTER TABLE  `distinta_base` ADD  `idmag_acq` INT( 11 ) ")
	conn.execute ("CREATE INDEX idmag_ven ON distinta_base (idmag_ven);")
	conn.execute ("CREATE INDEX idmag_acq ON distinta_base (idmag_acq);")
	call imposta_versione(nuova_versione)'207
end if

nuova_versione=209
if versione<nuova_versione  then
	sql="select magazzino_movimenti.id, magazzino.idmag from magazzino_movimenti inner join magazzino on (magazzino_movimenti.idpro=magazzino.idpro and magazzino_movimenti.idvara=magazzino.idvara and magazzino_movimenti.idvarb=magazzino.idvarb)"
	set rs=conn.Execute(sql)
	do while not rs.EOF
		conn.Execute("update magazzino_movimenti set idmag="&rs("idmag")&" where id="&rs("id"))
		rs.MoveNext
	loop
	call imposta_versione(nuova_versione)'209
end if

nuova_versione=210
if versione<nuova_versione  then
	conn.Execute("delete from magazzino_movimenti where idmag is null")
	call imposta_versione(nuova_versione)'210
end if

nuova_versione=211
if versione<nuova_versione  then
	conn.Execute("update magazzino_movimenti set causale=3 where causale=6")
	call imposta_versione(nuova_versione)'211
end if

nuova_versione=212
if versione<nuova_versione  then
	sql="select magazzino_movimenti.id, magazzino.idmag from magazzino_movimenti inner join magazzino on (magazzino_movimenti.idpro=magazzino.idpro and magazzino_movimenti.idvara=magazzino.idvara and magazzino_movimenti.idvarb=magazzino.idvarb)"
	set rs=conn.Execute(sql)
	do while not rs.EOF
		conn.Execute("update magazzino_movimenti set idmag="&rs("idmag")&" where id="&rs("id"))
		rs.MoveNext
	loop
	call imposta_versione(nuova_versione)'212
end if

nuova_versione=213
if versione<nuova_versione  then
	sql="select distinta_base.iddb, magazzino.idmag from distinta_base inner join magazzino on (distinta_base.idpro_ven=magazzino.idpro and distinta_base.idvara_ven=magazzino.idvara and distinta_base.idvarb_ven=magazzino.idvarb)"
	set rs=conn.Execute(sql)
	do while not rs.EOF
		conn.Execute("update distinta_base set idmag_ven="&rs("idmag")&" where iddb="&rs("iddb"))
		rs.MoveNext
	loop
	sql="select distinta_base.iddb, magazzino.idmag from distinta_base inner join magazzino on (distinta_base.idpro_acq=magazzino.idpro and distinta_base.idvara_acq=magazzino.idvara and distinta_base.idvarb_acq=magazzino.idvarb)"
	set rs=conn.Execute(sql)
	do while not rs.EOF
		conn.Execute("update distinta_base set idmag_acq="&rs("idmag")&" where iddb="&rs("iddb"))
		rs.MoveNext
	loop
	call imposta_versione(nuova_versione)'213
end if

nuova_versione=214
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_clienti` ADD  `stato_estero` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'214
end if

nuova_versione=215
if versione<nuova_versione  then
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql="select utenti.iduser from utenti where regione=21"
	set rs=conn.Execute(sql)
	do while not rs.EOF
		rs2.open "select iduser,stato_estero from utenti_clienti where iduser="&rs("iduser"),conn,3,3
		if rs2.eof then
			rs2.addnew
			rs2("iduser")=rs("iduser")
		end if
		rs2("stato_estero")=1
		rs2.update
		rs2.Close
		rs.MoveNext
	loop
	call imposta_versione(nuova_versione)'215
end if

nuova_versione=216
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `varianti_ordine` VARCHAR( 100 )  NULL DEFAULT  '' after articolo_ordine")
	conn.execute ("ALTER TABLE  `preventivi_dett` ADD  `varianti_ordine` VARCHAR( 100 )  NULL DEFAULT  '' after articolo_ordine")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett` ADD  `varianti_ordine` VARCHAR( 100 )  NULL DEFAULT  '' after articolo_ordine")
	call imposta_versione(nuova_versione)'216
end if

nuova_versione=217
if versione<nuova_versione  then
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	rs2.open "select ordini_dett.* from ordini_dett",conn,3,3
	do while not rs2.eof
		txt=""
		if rs2("var1")<>"" then
			txt = rs2("variante1_ordine") & ": " & rs2("var1") 
		end if
		if rs2("var2")<>"" then
			if txt<>"" then txt=txt&" - "
			txt =txt & rs2("variante2_ordine") & ": " & rs2("var2")
		end if		
		if txt<>"" then
			rs2("varianti_ordine")=txt
			rs2.update
		end if
					
		rs2.MoveNext
	loop	
	rs2.Close
	rs2.open "select preventivi_dett.* from preventivi_dett",conn,3,3
	do while not rs2.eof
		txt=""
		if rs2("var1")<>"" then
			txt = rs2("variante1_ordine") & ": " & rs2("var1") 
		end if
		if rs2("var2")<>"" then
			if txt<>"" then txt=txt&" - "
			txt =txt & rs2("variante2_ordine") & ": " & rs2("var2")
		end if		
		if txt<>"" then
			rs2("varianti_ordine")=txt
			rs2.update
		end if
					
		rs2.MoveNext
	loop	
	rs2.Close
	rs2.open "select ordini_fornitori_dett.* from ordini_fornitori_dett",conn,3,3
	do while not rs2.eof
		txt=""
		if rs2("var1")<>"" then
			txt = rs2("variante1_ordine") & ": " & rs2("var1") 
		end if
		if rs2("var2")<>"" then
			if txt<>"" then txt=txt&" - "
			txt =txt & rs2("variante2_ordine") & ": " & rs2("var2")
		end if		
		if txt<>"" then
			rs2("varianti_ordine")=txt
			rs2.update
		end if
					
		rs2.MoveNext
	loop	
	rs2.Close
	set rs2 = Nothing	
	call imposta_versione(nuova_versione)'217
end if

nuova_versione=218
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `quantita_scalato_magazzino` DECIMAL( 10, 3 ) NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'168
end if

nuova_versione=219
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` change  `quantita_scalato_magazzino` `quantita_scalato_magazzino` INT( 11 ) NOT NULL DEFAULT  '0' after quantita")
	call imposta_versione(nuova_versione)'219
end if
	
nuova_versione=220
if versione<nuova_versione  then
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql="select ordini.idord from ordini where stato=6"
	set rs=conn.Execute(sql)
	do while not rs.EOF
		conn.execute ("update ordini_dett set quantita_scalato_magazzino=quantita where idord="&rs("idord"))
		rs.movenext
	loop
	call imposta_versione(nuova_versione)'220
end if

nuova_versione=222
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_dett_fornitori_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_dett_fornitori_spettanze` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_fornitori_spettanze` ADD  `iddip` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_fornitori_spettanze` ADD  `quantita` INT( 3 ) NOT NULL")
	call imposta_versione(nuova_versione)'222
end if

nuova_versione=223
if versione<nuova_versione  then
	'conn.execute ("DROP TABLE IF EXISTS prodotti_ebay;")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `spettanze` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'223
end if

nuova_versione=224
if versione<nuova_versione  then
	conn.execute ("DROP TABLE IF EXISTS ordini_dett_fornitori_spettanze;")
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_fornitori_dett_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_spettanze` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_spettanze` ADD  `iddip` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_spettanze` ADD  `quantita` INT( 3 ) NOT NULL")
	call imposta_versione(nuova_versione)'224
end if

nuova_versione=226
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `quantita_ordinato` DECIMAL( 10, 3 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `quantita_impegnato` DECIMAL( 10, 3 ) NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'168
end if

nuova_versione=227
if versione<nuova_versione  then
	conn.execute ("DROP TABLE IF EXISTS ordini_dett_ordini_fornitori_dett;")
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_dett_ordini_fornitori_dett` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_dett_ordini_fornitori_dett` ADD  `iddett_ordini` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_ordini_fornitori_dett` ADD  `iddett_fornitori` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_ordini_fornitori_dett` ADD  `quantita` INT( 3 ) NOT NULL")
	call imposta_versione(nuova_versione)'227
end if

nuova_versione=228
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `admin` ADD  `permessi` VARCHAR( 100 ) NOT NULL DEFAULT ''")
	'memorizzo permessi in sessione, verifico con instr
	call imposta_versione(nuova_versione)'228
end if

nuova_versione=229
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `magazzino` ADD  `quantita_impegnata` FLOAT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'229
end if

nuova_versione=230
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `note_su_utente`  `note_su_utente` VARCHAR( 1000 ) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT  ''")
	conn.execute ("ALTER TABLE  `utenti_clienti` CHANGE  `FatturaPACodiceDestinatario`  `FatturaPACodiceDestinatario` VARCHAR( 500 ) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT  ''")
	conn.execute ("ALTER TABLE  `utenti_clienti` CHANGE  `email_inoltro_fattura`  `email_inoltro_fattura` VARCHAR( 500 ) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT  ''")
	call imposta_versione(nuova_versione)'230
end if

nuova_versione=231
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` ADD  `telefonare` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("update prodotti set telefonare=1 where prezzo=-1")
	call imposta_versione(nuova_versione)'231
end if

nuova_versione=232
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `numeri_di_serie`  int( 2 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'232
end if

nuova_versione=233
if versione<nuova_versione  then
	conn.execute ("update  ordini_dett set  numeri_di_serie  = richiedi_seriale")
	call imposta_versione(nuova_versione)'233
end if

nuova_versione=234
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `mepa_categorie` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `mepa_categorie` ADD  `descrizione` VARCHAR( 200 )")
	conn.execute ("ALTER TABLE  `mepa_categorie` ADD  `campi_JSON` VARCHAR( 20000 )")
	call imposta_versione(nuova_versione)'234
end if

nuova_versione=235
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `mepa_campi` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `idcategoria` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `label` VARCHAR( 100 ) NOT NULL")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `obbligatorio` VARCHAR( 1 )")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `digitabile` VARCHAR( 1 )")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `formato` VARCHAR( 100 )")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `dimensione` VARCHAR( 100 )")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `valori_default` VARCHAR( 500 )")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `range` VARCHAR( 100 )")
	conn.execute ("ALTER TABLE  `mepa_campi` ADD  `algoritmo` VARCHAR( 100 )")
	call imposta_versione(nuova_versione)'235
end if

nuova_versione=236
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `prodotti_mepa_campi` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `prodotti_mepa_campi` ADD  `idcategoria` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_mepa_campi` ADD  `idpro` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `prodotti_mepa_campi` ADD  `mepa_campi` TEXT NOT NULL")
	call imposta_versione(nuova_versione)'236
end if

nuova_versione=237
if versione<nuova_versione  then
	conn.execute ("DROP  TABLE IF EXISTS  `settori_settori`")
	conn.execute ("CREATE TABLE IF NOT EXISTS `settori_settori` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `settori_settori` ADD  `idpadre` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `settori_settori` ADD  `idfiglio` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `settori_settori` ADD  `ordine` INT( 11 ) NOT NULL default '0'")
	call imposta_versione(nuova_versione)'237
end if

nuova_versione=238
if versione<nuova_versione  then
	sql="select settori.* from settori "
	set rs=conn.Execute(sql)
	do while not rs.EOF
		ordine=rs("ordine")
		if isnull(ordine) then ordine = 0
		conn.execute ("INSERT INTO settori_settori (idpadre, idfiglio, ordine) VALUES ("&rs("idpadre")&", "&rs("idsettore")&", "&ordine&")")
		rs.movenext
	loop
	call imposta_versione(nuova_versione)'238
end if

nuova_versione=239
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX idpadre ON settori_settori (idpadre);")
	conn.execute ("CREATE INDEX idfiglio ON settori_settori (idfiglio); ")
	call imposta_versione(nuova_versione)'239
end if

nuova_versione=240
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX ordine ON ordini_dett (ordine ASC) ;")
	conn.execute ("CREATE INDEX ordine ON ordini_fornitori_dett (ordine ASC) ;")
	conn.execute ("CREATE INDEX ordine ON preventivi_dett (ordine ASC) ;")
	call imposta_versione(nuova_versione)'240
end if

nuova_versione=241
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX nord ON ordini (nord DESC) ;")
	conn.execute ("CREATE INDEX nord ON ordini_fornitori (nord DESC) ;")
	conn.execute ("CREATE INDEX nord ON preventivi (nord DESC) ;")
	conn.execute ("CREATE INDEX stato ON preventivi (stato) ;")
	conn.execute ("CREATE INDEX data ON preventivi (data) ;")
	conn.execute ("CREATE INDEX stato ON ordini_fornitori (stato) ;")
	conn.execute ("CREATE INDEX data ON ordini_fornitori (data) ;")
	call imposta_versione(nuova_versione)'241
end if

nuova_versione=242
if versione<nuova_versione  then
	on error resume next
	conn.execute ("ALTER TABLE  `settori` DROP COLUMN  `idtmp`")
	on error goto 0
	conn.execute ("ALTER TABLE  `settori` ADD  `prodotti_virtuali` INT( 2 ) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `settori` ADD  `idpro_virtuali` VARCHAR( 20000 ) default ''")
	call imposta_versione(nuova_versione)'242
end if

nuova_versione=243
if versione<nuova_versione  then
	on error resume next
	conn.execute ("ALTER TABLE  `settori` ADD  `propaga` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	err.clear
	on error goto 0	
	call imposta_versione(nuova_versione)'243
end if

nuova_versione=244
if versione<nuova_versione  then
	on error resume next
	conn.execute ("ALTER TABLE  `settori` ADD  `prendi_prodotti` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `settori` ADD  `cedi_prodotti` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'244
end if

nuova_versione=245
if versione<nuova_versione  then
	on error resume next
	conn.execute ("ALTER TABLE  `settori` ADD  `metodo` int( 1 )  NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'245
end if

nuova_versione=246
if versione<nuova_versione  then
	on error resume next
	conn.execute ("update ordini set tipopagamento=0 where tipopagamento=99")
	call imposta_versione(nuova_versione)'246
end if

nuova_versione=247
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` ADD  `prezzo_visibile` int( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("update prodotti set prezzo_visibile=1 where telefonare=1")
	conn.execute ("update prodotti set prezzo_visibile=2 where prezzo_riservato=1")
	call imposta_versione(nuova_versione)'247
end if

nuova_versione=248
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `email_news` ADD  `data_rimosso` datetime")
	call imposta_versione(nuova_versione)'248
end if

nuova_versione=249
if versione<nuova_versione  then
	conn.execute ("update  `ordini_dett` set idvara=0 WHERE `idpro` > 0 AND `idvara` IS NULL")
	conn.execute ("update  `ordini_dett` set idvarb=0 WHERE `idpro` > 0 AND `idvarb` IS NULL")
	call imposta_versione(nuova_versione)'249
end if

nuova_versione=250
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` DROP COLUMN  `noprezzo`")
	call imposta_versione(nuova_versione)'250
end if

nuova_versione=251
if versione<nuova_versione  then
	on error Resume next
	conn.execute ("ALTER TABLE settori_prodotti DROP INDEX id")
	err.clear
	on error goto 0
	conn.execute ("CREATE INDEX idpro ON prodotti_mepa_campi (idpro); ")
	conn.execute ("CREATE UNIQUE INDEX idpro_idcategoria ON prodotti_mepa_campi (idpro,idcategoria); ")
	call imposta_versione(nuova_versione)'251
end if

nuova_versione=252
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `articolo`  `articolo` VARCHAR( 100 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL; ")
	call imposta_versione(nuova_versione)'252
end if

nuova_versione=253
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD   `costo_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `costo_ordine`  `costo_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'253
end if

nuova_versione=254
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi_dett` ADD   `costo_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'254
end if

nuova_versione=255
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD   `costo_ordine` DECIMAL( 10, 2 ) NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'255
end if

nuova_versione=256
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD   `matricola` VARCHAR( 10 ) default ''")
	call imposta_versione(nuova_versione)'256
end if

nuova_versione=257
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_cosce` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_cavallo` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_polpacci` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `maglieria` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_tecnici` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_tecnici` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `prodotti` ADD  `tipo_taglia` INT( 2 ) NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'257
end if

nuova_versione=258
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti_spettanze`  DROP COLUMN  `tipo_prodotto` ")
	call imposta_versione(nuova_versione)'258
end if

nuova_versione=259
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `note` VARCHAR( 50 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_spettanze` ADD  `note` VARCHAR( 50 ) NULL DEFAULT ''")
	conn.execute ("CREATE TABLE IF NOT EXISTS `preventivi_dett_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `iddip` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `quantita` INT( 3 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `note` VARCHAR( 50 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'259
end if

nuova_versione=260
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `noti_email_registr`  VARCHAR( 255 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'260
end if

nuova_versione=261
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `noti_email_ordini`  VARCHAR( 255 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'261
end if

nuova_versione=262
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `taglia_misura` VARCHAR( 255 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'262
end if

nuova_versione=263
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `taglia_misura` VARCHAR( 255 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'263
end if

nuova_versione=264
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_clienti` ADD  `mod_dip` VARCHAR( 1000 ) NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'264
end if

nuova_versione=266
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_clienti`  DROP COLUMN  `mod_dip` ")
	conn.execute ("ALTER TABLE  `utenti_clienti` ADD  `mod_dip` VARCHAR( 1000 ) NOT NULL DEFAULT  ''")
	call imposta_versione(nuova_versione)'266
end if

nuova_versione=267
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `spettanze` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	call imposta_versione(nuova_versione)'267
end if

nuova_versione=268
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD   `arma` VARCHAR( 50 ) default ''")
	call imposta_versione(nuova_versione)'268
end if

nuova_versione=269
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_spettanze` ADD  `taglia_misura` VARCHAR( 255 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'269
end if

nuova_versione=270
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `gonna` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `cappotto` VARCHAR( 10 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'270
end if

nuova_versione=271
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `impegno_spesa` VARCHAR( 50 ) NULL DEFAULT '' after mepa")
	call imposta_versione(nuova_versione)'271
end if

nuova_versione=272
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `camicia_note` VARCHAR( 100 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_fondo2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_manica2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_torace2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_vita2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_spalle2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `cappotto2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_lunghezza2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_vita2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_bacino2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_cosce2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_cavallo2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_polpacci2` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `gonna2` varchar(10) DEFAULT ''")
	call imposta_versione(nuova_versione)'272
end if

nuova_versione=273
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_TG2` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `pantaloni_TG2` VARCHAR( 10 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'273
end if

nuova_versione=274
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `mepa_tipo` int( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after cig")
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `mepa`  `mepa_testo` VARCHAR( 50 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'274
end if

nuova_versione=275
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `mepa_data` DATETIME NULL DEFAULT NULL after mepa_testo")
	call imposta_versione(nuova_versione)'275
end if

nuova_versione=276
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE `ordini_dett` CHANGE `articolo_ordine` `articolo_ordine` TEXT CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE `preventivi_dett` CHANGE `articolo_ordine` `articolo_ordine` TEXT CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE `ordini_fornitori_dett` CHANGE `articolo_ordine` `articolo_ordine` TEXT CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'276
end if

nuova_versione=277
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `cig` VARCHAR( 12 ) NULL DEFAULT '' after d_cap")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `mepa_tipo` int( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after cig")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `mepa_testo` VARCHAR( 50 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT '' after mepa_tipo")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `mepa_data` DATETIME NULL DEFAULT NULL after mepa_testo")
	call imposta_versione(nuova_versione)'277
end if

nuova_versione=278
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `impegno_spesa` VARCHAR( 50 ) NULL DEFAULT '' after mepa_tipo")
	call imposta_versione(nuova_versione)'278
end if

nuova_versione=279
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `admin` CHANGE  `permessi`  `permessi` TINYTEXT CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT  ''")
	call imposta_versione(nuova_versione)'279
end if

nuova_versione=280
if versione<nuova_versione  then
	conn.execute ("update admin set  permessi='A1=1,A2=1,A3=1,A4=1,A5=1,B1=1,B2=1,C1=1,C2=1,C3=1,C4=1,D1=1,D2=1,E1=1,E2=1,E3=1,E4=1,F1=1,F2=1'")
	permessi=array("A1","A2","A3","A4","A5","B1","B2","C1","C2","C3","C4","D1","D2","E1","E2","E3","E4","F1","F2","Z1","Z2")
	call imposta_versione(nuova_versione)'280
end if


nuova_versione=281
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` ADD    `selezmailing` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after mailing")
	call imposta_versione(nuova_versione)'281
end if

nuova_versione=282
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_bacino` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `giacca_bacino2` varchar(10) DEFAULT ''")
	call imposta_versione(nuova_versione)'282
end if

nuova_versione=283
if versione<nuova_versione  then
	conn.execute("INSERT INTO `traduzioni` ( `pagina`, `chiave`, `valore`, `valore_en`) VALUES( 'tutte', 'preventiv', 'PREVENTIVO', 'PREVENTIVO');")
	call imposta_versione(nuova_versione)'283
end if

nuova_versione=284
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_dett_spettanze_dettaglio` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD  `idspettanza` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD  `quantita` INT( 11 ) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD  `data` DATETIME NOT NULL")
	call imposta_versione(nuova_versione)'284
end if

nuova_versione=285
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD  `iddett` INT( 11 ) NOT NULL after id")
	call imposta_versione(nuova_versione)'285
end if

nuova_versione=286
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD  `stato` INT( 2 ) NOT NULL after idspettanza")
	call imposta_versione(nuova_versione)'286
end if

nuova_versione=287
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD INDEX (  `iddett` )")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_spettanze` ADD INDEX (  `iddett` )")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD INDEX (  `iddett` )")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD INDEX (  `iduser` )")
	on error resume next
	conn.execute ("ALTER TABLE ordini_dett_originale DROP INDEX iddett")
	err.clear
	on error goto 0
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD INDEX (  `iduser` )")
	conn.execute ("ALTER TABLE  `prodotti_spettanze` ADD INDEX (  `idpro` ,  `idvara` ,  `idvarb` ) ;")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD INDEX (  `iddip` )")
	call imposta_versione(nuova_versione)'287
end if

nuova_versione=288
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD INDEX (  `iddett` )")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze_dettaglio` ADD INDEX (  `idspettanza` )")
	call imposta_versione(nuova_versione)'288
end if

nuova_versione=289
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `in_ordine` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `pronto` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini_dett_spettanze` ADD  `consegnato` int(3) not NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'289
end if

nuova_versione=290
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `in_ordine` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `pronto` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini_dett` ADD  `consegnato` int(3) not NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'290
end if

nuova_versione=291
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `impegno_spesa`  `impegno_spesa` VARCHAR( 200 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'291
end if

nuova_versione=292
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_dett_note` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_dett_note` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_dett_note` ADD  `nota`  VARCHAR( 1000 ) not null default ''")
	call imposta_versione(nuova_versione)'292
end if

nuova_versione=293
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_fornitori_dett_note` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_note` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_note` ADD  `nota`  VARCHAR( 1000 ) not null default ''")
	conn.execute ("CREATE TABLE IF NOT EXISTS `preventivi_dett_note` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `preventivi_dett_note` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `preventivi_dett_note` ADD  `nota`  VARCHAR( 1000 ) not null default ''")
	call imposta_versione(nuova_versione)'293
end if

nuova_versione=294
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `cintura` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `cinturone` varchar(10) DEFAULT ''")
	conn.execute ("ALTER TABLE `ordini_dett_note` DROP `id`;")
	conn.execute ("ALTER TABLE `ordini_fornitori_dett_note` DROP `id`;")
	conn.execute ("ALTER TABLE `preventivi_dett_note` DROP `id`;")
	conn.execute ("ALTER TABLE  `ordini_dett_note` ADD PRIMARY KEY (  `iddett` ) ;")
	conn.execute ("ALTER TABLE  `ordini_fornitori_dett_note` ADD PRIMARY KEY (  `iddett` ) ;")
	conn.execute ("ALTER TABLE  `preventivi_dett_note` ADD PRIMARY KEY (  `iddett` ) ;")
	conn.execute ("CREATE INDEX data_rilievo ON dipendenti_misure (data_rilievo);")
	conn.execute ("CREATE INDEX ricerca ON files (cosa,idcosa);")
	call imposta_versione(nuova_versione)'294
end if

nuova_versione=295
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `ornamento` VARCHAR( 100 ) NOT NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'295
end if

nuova_versione=296
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ddt_dett` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ddt_dett` ADD  `idddt` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ddt_dett` ADD  `iddett` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `ddt_dett` ADD  `quantita` INT( 11 ) NOT NULL default '0'")
	conn.execute ("CREATE INDEX idddt ON ddt_dett (idddt);")
	conn.execute ("ALTER TABLE  `ddt` ADD  `parziale` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'296
end if

nuova_versione=297
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `nascondi_totali` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'297
end if

nuova_versione=298
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `parziale` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'298
end if

nuova_versione=299
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `scadenze_for` CHANGE  `importo`  `importo` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `fatture_for` CHANGE  `totale_fattura`  `totale_fattura` DECIMAL( 10, 2 ) NOT NULL DEFAULT  '0'")

	call imposta_versione(nuova_versione)'299
end if
nuova_versione=300
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `messaggi_utenti` (`id` int(11) NOT NULL AUTO_INCREMENT,`idmessaggio` int(11) NOT NULL,`iduser` int(11) NOT NULL,`data_lettura` datetime DEFAULT NULL,PRIMARY KEY (`id`),KEY `idmessaggio` (`idmessaggio`),KEY `iduser` (`iduser`)) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=41 ;")
	conn.execute ("CREATE TABLE IF NOT EXISTS `messaggi_discussioni` (  `iddiscussione` int(11) NOT NULL AUTO_INCREMENT, `iduser` int(11) NOT NULL,  `data_apertura` datetime NOT NULL, `ultima_risposta` datetime DEFAULT NULL,`Oggetto` longtext NOT NULL,PRIMARY KEY (`iddiscussione`)) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=14 ;")
	conn.execute ("CREATE TABLE IF NOT EXISTS `messaggi` (  `idmessaggio` int(11) NOT NULL AUTO_INCREMENT,  `iduser` int(11) NOT NULL,  `iddiscussione` int(11) NOT NULL,  `data` datetime NOT NULL,  `messaggio` longtext NOT NULL,  PRIMARY KEY (`idmessaggio`)) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=18 ;")
	call imposta_versione(nuova_versione)'300
end if

nuova_versione=301
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `messaggi_discussioni` ADD  `destinatari` TEXT NOT NULL")

	call imposta_versione(nuova_versione)'301
end if

nuova_versione=302
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `discussioni_destinatari` (`iduser` int(11) NOT NULL,`iddiscussione` int(11) NOT NULL) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=41 ;")

	call imposta_versione(nuova_versione)'302
end if
nuova_versione=303
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `discussioni_destinatari` ADD UNIQUE ( `iduser` , `iddiscussione`);")

	call imposta_versione(nuova_versione)'303
end if

nuova_versione=304
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `messaggi_utenti` ADD  `iddiscussione` INT( 11 ) NOT NULL default '0'")

	call imposta_versione(nuova_versione)'304
end if

nuova_versione=305
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `messaggi_discussioni` ADD  `scadenza` datetime DEFAULT NULL")
	conn.execute ("ALTER TABLE  `messaggi_discussioni` ADD  `stato` int(1) DEFAULT '0'")
	call imposta_versione(nuova_versione)'304
end if

nuova_versione=306
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `varianti_a` ADD  `codiceVariante` VARCHAR( 100 ) NOT NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'306
end if
nuova_versione=307
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE `varianti_a` DROP `codicevariante`;")
	conn.execute ("ALTER TABLE `varianti_a` ADD  `codiceVarA` VARCHAR( 100 ) NOT NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'307
end if
nuova_versione=308
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` CHANGE  `creato_da_admin`  `creato_da_admin` int(11) NOT NULL DEFAULT  '0' after iduser")	
	conn.execute ("ALTER TABLE  `preventivi` CHANGE  `creato_da_admin`  `creato_da_admin` int(11) NOT NULL DEFAULT  '0' after iduser")	
	conn.execute ("ALTER TABLE  `ordini_fornitori` CHANGE  `creato_da_admin`  `creato_da_admin` int(11) NOT NULL DEFAULT  '0' after iduser")	
	call imposta_versione(nuova_versione)'308
end if
nuova_versione=309
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `cognome` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `nome` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "utenti_dipendenti",conn,3,3
	do while not rs.eof
		if rs("nominativo")<>"" then
			nominativov=replace(rs("nominativo"),"  "," ")
			arraytxt=split(nominativov)
			rs("cognome")=arraytxt(0)
			if ubound (arraytxt)>0 then
				rs("nome")=arraytxt(1)
				if ubound (arraytxt)>1 then
					rs("nome")=arraytxt(1)&" "&arraytxt(2)
				end if
			end if
			rs.update
		end if
	
	rs.MoveNext
	loop
	rs.Close
	set rs = Nothing
	call imposta_versione(nuova_versione)'309
end if
nuova_versione=310
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `lingue` VARCHAR( 80 ) NOT NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'310
end if
nuova_versione=311
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_dipendenti` change  `lingue` `lingue` VARCHAR( 80 )  DEFAULT ''")
	call imposta_versione(nuova_versione)'311
end if


nuova_versione=312
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `colore` VARCHAR( 8 )  DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `ordine` INT( 8 )  DEFAULT '0'")
	call imposta_versione(nuova_versione)'312
end if

nuova_versione=313
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `appunti` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `appunti` ADD  `iduser` int(1) DEFAULT '0'")
	conn.execute ("ALTER TABLE  `appunti` ADD  `pubblico` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `appunti` ADD  `data` datetime DEFAULT NULL")
	conn.execute ("ALTER TABLE  `appunti` ADD  `fatto` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `appunti` ADD  `stato` INT( 2 )  DEFAULT '0'")
	call imposta_versione(nuova_versione)'313
end if

nuova_versione=314
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `appunti` ADD  `testo` VARCHAR(1000)")
	call imposta_versione(nuova_versione)'313
end if

nuova_versione=315
if versione<nuova_versione  then
	conn.execute ("RENAME TABLE  `utenti_dipendenti` TO  `dipendenti`")
	call imposta_versione(nuova_versione)'315
end if

nuova_versione=316
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_dipendenti` ( `iddip` int(11) NOT NULL) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `iduser` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE utenti_dipendenti ADD PRIMARY KEY (iddip, iduser);")
	call imposta_versione(nuova_versione)'316
end if

nuova_versione=317
if versione<nuova_versione  then
	conn.execute ("delete from utenti_dipendenti")
	set rs=conn.execute("select * from dipendenti")
	do while not rs.EOF
		conn.execute ("insert into utenti_dipendenti (iddip,iduser) values ("&rs("iddip")&","&rs("iduser")&")")
		rs.MoveNext
	loop
	
	'call imposta_versione(nuova_versione)'317
end if
nuova_versione=318
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `appunti` CHANGE  `iduser` `iduser` int(11) DEFAULT '0'")
	conn.execute ("CREATE TABLE IF NOT EXISTS `appunti_letture` ( `idappunto` int(11) NOT NULL) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `appunti_letture` ADD  `iduser` int(11) DEFAULT '0'")
	conn.execute ("ALTER TABLE appunti_letture ADD PRIMARY KEY (idappunto, iduser);")
	call imposta_versione(nuova_versione)'318
end if

nuova_versione=319
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `anno`  year(4) NOT NULL default '0' after data")
	conn.execute("update ordini set anno=year(data)")
	conn.execute ("CREATE INDEX anno ON ordini (anno);")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `anno`  year(4) NOT NULL default '0'  after data")
	conn.execute("update preventivi set anno=year(data)")
	conn.execute ("CREATE INDEX anno ON preventivi (anno);")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `anno`  year(4) NOT NULL default '0'  after data")
	conn.execute("update ordini_fornitori set anno=year(data)")
	conn.execute ("CREATE INDEX anno ON ordini_fornitori (anno);")
	conn.execute ("ALTER TABLE  `fatture` ADD  `anno`  year(4) NOT NULL default '0'  after data")
	conn.execute("update fatture set anno=year(data)")
	conn.execute ("CREATE INDEX anno ON fatture (anno);")
	conn.execute ("ALTER TABLE  `ddt` ADD  `anno`  year(4) NOT NULL default '0'  after data")
	conn.execute("update ddt set anno=year(data)")
	conn.execute ("CREATE INDEX anno ON ddt (anno);")
	call imposta_versione(nuova_versione)'319
end if

nuova_versione=320
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `idagente`  int(11) NOT NULL default '0' after creato_da_admin")
	call imposta_versione(nuova_versione)'320
end if

nuova_versione=321
if versione<nuova_versione  then
	conn.execute("update ordini set creato_da_admin=-1 where creato_da_admin=1")
	call imposta_versione(nuova_versione)'321
end if

nuova_versione=322
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX creato_da_admin ON ordini (creato_da_admin);")
	conn.execute ("CREATE INDEX creato_da_admin ON preventivi (creato_da_admin);")
	call imposta_versione(nuova_versione)'322
end if

nuova_versione=323
if versione<nuova_versione  then
	conn.execute("update admin set permessi = CONCAT('C5=1,', permessi)")
	call imposta_versione(nuova_versione)'323
end if

nuova_versione=324
if versione<nuova_versione  then
	conn.execute("update preventivi set creato_da_admin=-1 where creato_da_admin=1")
	call imposta_versione(nuova_versione)'324
end if

nuova_versione=325
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX idagente ON ordini (idagente);")
	call imposta_versione(nuova_versione)'324
end if

nuova_versione=326
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ddt` ADD  `tipo` ENUM(  '0',  '1',  '2' ) NOT NULL DEFAULT  '0', ADD INDEX (  `tipo` )")
	conn.execute ("ALTER TABLE  `ddt` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `ddt` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON ddt (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON ddt (idconsegna);")
	
	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_intestazioni` (`id` int(11) NOT NULL AUTO_INCREMENT,`iduser` int(11) DEFAULT NULL,`Cognome` varchar(255) DEFAULT NULL,`Nome` varchar(255) DEFAULT NULL,`Azienda` varchar(65) DEFAULT NULL,`Indirizzo` varchar(50) DEFAULT NULL,`Citta` varchar(50) DEFAULT NULL,`cap` varchar(10) DEFAULT NULL,`Provincia` varchar(50) DEFAULT NULL,`Regione` int(11) DEFAULT NULL,`Piva` varchar(50) DEFAULT NULL,`CF` varchar(20) DEFAULT NULL,PRIMARY KEY (`id`),KEY `iduser` (`iduser`)) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;")

	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_consegna` (`id` int(11) NOT NULL AUTO_INCREMENT,`iduser` int(11) DEFAULT NULL,`D_Azienda` varchar(65) DEFAULT NULL,`D_Indirizzo` varchar(50) DEFAULT NULL,`D_Citta` varchar(50) DEFAULT NULL,`D_Regione` int(11) DEFAULT NULL,`D_provincia` varchar(50) DEFAULT NULL,`D_cap` varchar(50) DEFAULT NULL,PRIMARY KEY (`id`),KEY `iduser` (`iduser`)) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;")
	
	call imposta_versione(nuova_versione)'326
end if

nuova_versione=327
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `ordini` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON ordini (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON ordini (idconsegna);")

	conn.execute ("ALTER TABLE  `preventivi` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON preventivi (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON preventivi (idconsegna);")

	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON ordini_fornitori (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON ordini_fornitori (idconsegna);")
	
	conn.execute ("ALTER TABLE  `fatture` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `fatture` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON fatture (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON fatture (idconsegna);")
	call imposta_versione(nuova_versione)'327
end if

nuova_versione=328
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `ordini` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON ordini (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON ordini (idconsegna);")

	conn.execute ("ALTER TABLE  `preventivi` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `preventivi` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON preventivi (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON preventivi (idconsegna);")

	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON ordini_fornitori (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON ordini_fornitori (idconsegna);")
	
	conn.execute ("ALTER TABLE  `fatture` ADD  `idintestazione`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `fatture` ADD  `idconsegna`  int(11) NOT NULL default '0' ")
	conn.execute ("CREATE INDEX idintestazione ON fatture (idintestazione);")
	conn.execute ("CREATE INDEX idconsegna ON fatture (idconsegna);")
'Ripetizione 327 per punto break con ricami piemonte	
	
	conn.execute ("ALTER TABLE  `utenti` ADD  `idintestazione`  int(11) NOT NULL default '0' after iduser")
	conn.execute ("CREATE INDEX idintestazione ON utenti (idintestazione);")
	call imposta_versione(nuova_versione)'328
end if

nuova_versione=329
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_intestazioni` CHANGE  `Indirizzo`  `Indirizzo` VARCHAR( 255 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `utenti_intestazioni` CHANGE  `nome`  `nome` VARCHAR( 40 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `utenti_intestazioni` CHANGE  `cognome`  `cognome` VARCHAR( 40 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `utenti_intestazioni` CHANGE  `azienda`  `azienda` VARCHAR( 100 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `utenti_consegna` CHANGE  `D_Indirizzo`  `D_Indirizzo` VARCHAR( 255 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `utenti_consegna` CHANGE  `d_azienda`  `d_azienda` VARCHAR( 100 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'329
end if

nuova_versione=330
if versione<nuova_versione  then
	conn.execute ("RENAME TABLE  `ddt_dett` TO  `ddt_dett_ordini`")
	call imposta_versione(nuova_versione)'330
	
end if

nuova_versione=331
if versione<nuova_versione  then
	Server.ScriptTimeout=60*20
	call popola_utenti_intestazioni()
	call imposta_versione(nuova_versione)'331
end if

nuova_versione=332
if versione<nuova_versione  then
	Server.ScriptTimeout=60*20
	call popola_utenti_consegna()
	call imposta_versione(nuova_versione)'331
end if

nuova_versione=333
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `tipo_documento` ENUM( '', 'ordine',  'fornitore',  'preventivo', 'ddt', 'fattura' )  NOT NULL DEFAULT  '' after idord, ADD INDEX (  `tipo_documento` )")
	conn.execute("update ordini set tipo_documento='ordine'")
	conn.execute ("ALTER TABLE  `ddt` ADD  `sub_idord`  int(11) NOT NULL default '0'")
	call imposta_versione(nuova_versione)'332
end if

nuova_versione=334
if versione<nuova_versione  then
	Server.ScriptTimeout=60*10
	call migra_ddt()
	call imposta_versione(nuova_versione)'334
end if

nuova_versione=335
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  nome")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  Cognome")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  azienda")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  cap")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  citta")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  indirizzo")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  regione")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  provincia")
	call imposta_versione(nuova_versione)'335
end if

nuova_versione=336
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ddt` DROP COLUMN  idconsegna")
	conn.execute ("ALTER TABLE  `ddt` DROP COLUMN  idintestazione")
	call imposta_versione(nuova_versione)'336
end if

nuova_versione=337
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ddt` DROP COLUMN  nddt, DROP COLUMN  data, DROP COLUMN  anno, drop iduser, drop destinatario, drop destinazione, drop eliminato")
	
	call imposta_versione(nuova_versione)'337
end if

nuova_versione=338
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `dati_mepa` ( `id` int(11) NOT NULL) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `dati_mepa` ADD  `cig` VARCHAR( 12 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dati_mepa` ADD  `mepa_tipo` int( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	conn.execute ("ALTER TABLE  `dati_mepa` ADD  `mepa_testo` VARCHAR( 50 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT '' ")
	conn.execute ("ALTER TABLE  `dati_mepa` ADD  `mepa_data` DATETIME NULL DEFAULT NULL ")
	conn.execute ("ALTER TABLE  `dati_mepa` ADD  `impegno_spesa` VARCHAR( 50 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'338
end if



nuova_versione=339
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `rollback` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `rollback` ADD  `idord`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `rollback` ADD  `data` datetime NOT NULL")
	conn.execute ("ALTER TABLE  `rollback` ADD  `descrizione` VARCHAR( 50 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `rollback` ADD  `eseguito` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	
	
	conn.execute ("CREATE TABLE IF NOT EXISTS `rollback_ordini_dett` ( `id` int(11) NOT NULL) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett` ADD  `idrollback`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett` ADD  `iddett`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett` ADD  `in_ordine` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett` ADD  `pronto` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett` ADD  `consegnato` int(3) not NULL DEFAULT '0'")
	
	conn.execute ("CREATE TABLE IF NOT EXISTS `rollback_ordini_dett_spettanze` ( `id` int(11) NOT NULL) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett_spettanze` ADD  `idrollback`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett_spettanze` ADD  `idspett`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett_spettanze` ADD  `in_ordine` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett_spettanze` ADD  `pronto` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `rollback_ordini_dett_spettanze` ADD  `consegnato` int(3) not NULL DEFAULT '0'")
	
	call imposta_versione(nuova_versione)'339
end if

nuova_versione=340
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ddt` ADD  `idrollback`  int(11) NOT NULL default '0'")
	call imposta_versione(nuova_versione)'340
end if

nuova_versione=341
if versione<nuova_versione  then
	'conn.execute ("UPDATE ddt set sub_idord=idord where sub_idord=ddt_dett_ordini")
	call imposta_versione(nuova_versione)'341
end if

nuova_versione=342
if versione<nuova_versione  then
	call migra_ddt2()
	call imposta_versione(nuova_versione)'342
end if
nuova_versione=343
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  fatture ADD  tipo_fattura ENUM( 'ordine', 'ddt' )  NOT NULL DEFAULT  'ordine' after idfat")
	call imposta_versione(nuova_versione)'343
end if

nuova_versione=344
if versione<nuova_versione  then
	call modifica_ordini_fatture()
	call imposta_tipo_fattura()
	
	call imposta_versione(nuova_versione)'344
end if

nuova_versione=345
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` ADD  `nascosto`  TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after attivo")
	
	call imposta_versione(nuova_versione)'345
end if

nuova_versione=346
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE `ordini` CHANGE `tipo_documento` `tipo_documento` ENUM('/','ordine','fornitore','preventivo','ddt','fattura','notacredito') CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT '/'")
	
	call imposta_versione(nuova_versione)'346
end if

nuova_versione=347
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE `scadenze` ADD `pagato` int( 1 ) UNSIGNED NOT NULL DEFAULT  '0'")
	
	call imposta_versione(nuova_versione)'347
end if

nuova_versione=348
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `richiedi_banca` TINYINT (1) UNSIGNED NOT NULL  DEFAULT  '0' after riba")
	conn.execute ("ALTER TABLE  `metodi_pagamento` ADD  `pagamentoPA` char (4)   DEFAULT  '' ")
	
	call imposta_versione(nuova_versione)'348
end if
nuova_versione=349
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti` ADD  `lato_arma` ENUM(  '',  'dx',  'sx' ) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT  ''")

	call imposta_versione(nuova_versione)'349
end if
nuova_versione=350
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `fatture_for` CHANGE  `nfat`  `nfat` VARCHAR( 50 ) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL")

	call imposta_versione(nuova_versione)'350
end if
nuova_versione=351
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ddt` ADD  `da_fatturare`   TINYINT (1) UNSIGNED NOT NULL  DEFAULT  '0' after causale")
	conn.execute ("UPDATE ddt set da_fatturare=1 where causale=1")
	call imposta_versione(nuova_versione)'351
end if
nuova_versione=352
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini_dett` CHANGE  `quantita`  `quantita` DECIMAL( 10, 2 ) NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `ddt_dett_ordini` CHANGE  `quantita`  `quantita` DECIMAL( 10, 2 ) NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'352
end if


nuova_versione=353
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE `ordini` CHANGE `tipo_documento` `tipo_documento` ENUM('/','ordine','fornitore','preventivo','ddt','fattura','notacredito','sostituzione') CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT '/'")
	
	call imposta_versione(nuova_versione)'353
end if

nuova_versione=354
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  ddt ADD  tipo_ddt ENUM( 'completo', 'parziale','no_ordine' )  NOT NULL DEFAULT  'completo' after idord")
	call imposta_versione(nuova_versione)'354
end if

nuova_versione=355
if versione<nuova_versione  then
	conn.execute ("update ddt set tipo_ddt='parziale' where tipo='1'")
	conn.execute ("update ddt set tipo_ddt='no_ordine' where tipo='2'")
	call imposta_versione(nuova_versione)'355
end if

nuova_versione=356
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ddt` DROP COLUMN  tipo")
	
	call imposta_versione(nuova_versione)'356
end if


nuova_versione=357
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `consenti_acquisti`  TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("UPDATE impostazioni set consenti_acquisti=1")
	call imposta_versione(nuova_versione)'357
end if


nuova_versione=358
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `tags` ADD  `testo` TEXT")
	conn.execute ("ALTER TABLE  `tags` ADD  `tipo` INT(1) UNSIGNED")
	call imposta_versione(nuova_versione)'358
end if
nuova_versione=359
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `tags` DROP  COLUMN `idnonno`")
	conn.execute ("ALTER TABLE  `tags` DROP  COLUMN `NomeConParenti`")
	conn.execute ("ALTER TABLE  `tags` DROP  COLUMN `idpadre`")
	call imposta_versione(nuova_versione)'359
end if
nuova_versione=360
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `colora_nominativo` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'360
end if
nuova_versione=361
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_gradi` ADD  `colore_nominativo` VARCHAR( 8 )  DEFAULT ''")
	conn.execute ("ALTER TABLE  `utenti_gradi` drop column  `colora_nominativo`")

	call imposta_versione(nuova_versione)'361
end if
nuova_versione=362
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `impostazioni` ADD  `consenti_registrazione`  TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '1'")
	call imposta_versione(nuova_versione)'362
end if

nuova_versione=363
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX sub_idord ON ddt (sub_idord);")
	call imposta_versione(nuova_versione)'363
end if


nuova_versione=364
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `in_ordine` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `pronto` int(3) not NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `preventivi_dett_spettanze` ADD  `consegnato` int(3) not NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'364
end if


nuova_versione=365
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti` ADD  `tipo_rilievo` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")

	call imposta_versione(nuova_versione)'365
end if

nuova_versione=366
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti` DROP COLUMN  `tipo_rilievo`")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `tipo_rilievo` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'366
end if

nuova_versione=367
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` CHANGE  `articolo`  `articolo` VARCHAR( 150 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL; ")
	call imposta_versione(nuova_versione)'367
end if

nuova_versione=368
if versione<nuova_versione  then
	conn.execute ("update admin set  permessi=concat(permessi,',A6=1')")
	call imposta_versione(nuova_versione)'368
end if

nuova_versione=369
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `agenti` ADD  `iduser` int(1) DEFAULT '0'")
	conn.execute ("CREATE INDEX iduser ON agenti (iduser);")
	
	call imposta_versione(nuova_versione)'369
end if

nuova_versione=370
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `provvigioni_agente` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `provvigioni_agente` ADD  `idagente`  int(11) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `provvigioni_agente` ADD  `sconto`   DECIMAL( 10, 2 ) NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `provvigioni_agente` ADD  `provvigione`   DECIMAL( 10, 2 ) NULL DEFAULT NULL")
	conn.execute ("CREATE INDEX idagente ON provvigioni_agente (idagente);")
	call imposta_versione(nuova_versione)'370
end if


nuova_versione=371
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` ADD  `visibilita`  int( 1 ) UNSIGNED NOT NULL DEFAULT  '0' after nascosto")
	call imposta_versione(nuova_versione)'371
end if

nuova_versione=372
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` DROP COLUMN  `vetrina`")
	conn.execute ("ALTER TABLE  `prodotti` DROP COLUMN  `click`")
	call imposta_versione(nuova_versione)'366
end if

nuova_versione=373
if versione<nuova_versione  then
	conn.execute ("update admin set  permessi=concat(permessi,',E5=1')")
	call imposta_versione(nuova_versione)'373
end if

nuova_versione=374
if versione<nuova_versione  then
	conn.execute ("CREATE INDEX idagente ON utenti (idagente);")
	call imposta_versione(nuova_versione)'374
end if
nuova_versione=375
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `idagente`  int(11) NOT NULL default '0' after creato_da_admin")
	conn.execute ("CREATE INDEX idagente ON preventivi (idagente);")
	call imposta_versione(nuova_versione)'375
end if

nuova_versione=376
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `ordini_sostituzioni` ( `idord` int(11) NOT NULL) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `ordini_sostituzioni` ADD  `idsostituzione`  int(11) NOT NULL")
	conn.execute("ALTER TABLE  `ordini_sostituzioni` ADD PRIMARY KEY (  `idord` ,  `idsostituzione` ) ;")
	call imposta_versione(nuova_versione)'376
end if
nuova_versione=377
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `prodotti` DROP COLUMN  nascosto")
	call imposta_versione(nuova_versione)'377
end if

nuova_versione=378
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `agenti` ADD  `clienti_provincie` VARCHAR( 10)  DEFAULT ''")
	conn.execute ("ALTER TABLE  `agenti` ADD  `clienti_pa` INT(1)  UNSIGNED NOT NULL DEFAULT  '0'")
	
	call imposta_versione(nuova_versione)'378
end if
nuova_versione=379
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_pantaloni` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_giacca` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_camicia` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_giaccatec` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_pantalonitec` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_cappotto` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_gonna` varchar(200) DEFAULT ''")
	
	call imposta_versione(nuova_versione)'379
end if

nuova_versione=380
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `ordini` ADD  `primo` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'380
end if

nuova_versione=381
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `appunti_destinatari` (`idappunto` int(11) NOT NULL,`iduser` int(11) NOT NULL,`data_lettura` datetime DEFAULT NULL,PRIMARY KEY (`idappunto`,`iduser`), KEY `data_lettura` (`data_lettura`)) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=41 ;")
	call imposta_versione(nuova_versione)'381
end if

nuova_versione=382
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti` DROP COLUMN  iduser")
	conn.execute ("ALTER TABLE  `dipendenti` ADD  `nuovo` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '1'")
	conn.execute ("update dipendenti set nuovo=0")
	
	call imposta_versione(nuova_versione)'382
end if

nuova_versione=383
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` ADD  `permessi_utente` VARCHAR( 50 )  DEFAULT ''")

	call imposta_versione(nuova_versione)'383
end if

nuova_versione=384
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti` ADD  `dummy1` TINYINT( 1 ) UNSIGNED DEFAULT '0' after iddip")

	call imposta_versione(nuova_versione)'384
end if

nuova_versione=385
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `appunti` CHANGE  `testo` `testo` TEXT")
	call imposta_versione(nuova_versione)'385
end if
nuova_versione=386
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  cf")
	conn.execute ("ALTER TABLE  `utenti` DROP COLUMN  piva")
	call imposta_versione(nuova_versione)'386
end if
nuova_versione=387
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_dipendenti` ADD  `dimesso` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	call imposta_versione(nuova_versione)'387
end if

nuova_versione=388
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `utenti_spettanze` ADD  `idpro` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `utenti_spettanze` ADD  `iddip` INT( 11 ) NOT NULL")
	conn.execute ("ALTER TABLE  `utenti_spettanze` ADD  `quantita` INT( 3 ) NOT NULL")
	conn.execute ("CREATE INDEX idpro ON utenti_spettanze (idpro);")
	conn.execute ("CREATE INDEX iddip ON utenti_spettanze (iddip);")

	call imposta_versione(nuova_versione)'388
end if
nuova_versione=389
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_spettanze` ADD  `iduser` INT( 11 ) NOT NULL")
	conn.execute ("CREATE INDEX iduser ON utenti_spettanze (iduser);")

	call imposta_versione(nuova_versione)'389
end if



nuova_versione=390
if versione<nuova_versione  then
	conn.execute ("CREATE TABLE IF NOT EXISTS `utenti_foglio_spettanze` ( `id` int(11) NOT NULL AUTO_INCREMENT, PRIMARY KEY (`id`)) ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=1;")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `iduser` INT( 11 ) NOT NULL")
	conn.execute ("CREATE INDEX iduser ON utenti_foglio_spettanze (iduser);")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `tipo_ordine` INT( 1 ) NOT NULL default '0'")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `cig` varchar(255) NULL default ''")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `determinazione` varchar(255) NULL default ''")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `note` text  NULL default ''")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD FOREIGN KEY (iduser) REFERENCES utenti(iduser) ON DELETE CASCADE")
	call imposta_versione(nuova_versione)'390
end if

nuova_versione=391
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `completo` datetime")
	call imposta_versione(nuova_versione)'391
end if

nuova_versione=392
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_spettanze` ADD  `idfoglio` INT( 11 ) NOT NULL")
	conn.execute ("CREATE INDEX idfoglio ON utenti_spettanze (idfoglio);")
	conn.execute ("ALTER TABLE  `utenti_spettanze` ADD FOREIGN KEY (idfoglio) REFERENCES utenti_foglio_spettanze(id) ON DELETE CASCADE")
	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `creato` datetime")

	call imposta_versione(nuova_versione)'392
end if
nuova_versione=393
if versione<nuova_versione  then

	conn.execute ("ALTER TABLE  `utenti_foglio_spettanze` ADD  `idord` INT( 11 )")

	call imposta_versione(nuova_versione)'393
end if
nuova_versione=394
if versione<nuova_versione  then

	conn.execute ("ALTER TABLE  `fatture` CHANGE  `tipo_fattura`  `tipo_fattura` ENUM(  'ordine',  'ddt',  'fattura' ) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT  'ordine'")

	call imposta_versione(nuova_versione)'394
end if


nuova_versione=395
if versione<nuova_versione  then
	conn.execute ("DROP TABLE IF EXISTS fatture_dett")
	conn.execute ("CREATE  TABLE  `fatture_dett` (  `iddett` int( 11  )  NOT  NULL  AUTO_INCREMENT ,`idfat` int( 11  )  NOT  NULL ,`idpro` int( 11  )  DEFAULT NULL , `var1` varchar( 50  )  DEFAULT NULL , `var2` varchar( 50  )  DEFAULT NULL , `um` varchar( 50  )  DEFAULT NULL , `quantita` decimal( 10, 2  )  DEFAULT NULL , `quantita_scalato_magazzino` int( 11  )  NOT  NULL DEFAULT  '0', `prezzo` decimal( 10, 2  )  NOT  NULL , `modificato` tinyint( 1  )  NOT  NULL , `sconto_prodotto` decimal( 10, 2  )  NOT  NULL , `codice_ordine` varchar( 70  )  DEFAULT NULL , `articolo_ordine` text, `varianti_ordine` varchar( 100  ) DEFAULT  '',`variante1_ordine` varchar( 50  )  DEFAULT NULL ,`variante2_ordine` varchar( 50  )  DEFAULT NULL , `reso` tinyint( 1  )  DEFAULT NULL ,`idvara` int( 11  )  DEFAULT NULL , `idvarb` int( 11  )  DEFAULT NULL , `ordine` int( 11  )  DEFAULT NULL , `costo_ordine` decimal( 10, 2  )  NOT  NULL DEFAULT  '0.00', `totale_riga` decimal( 10, 2  )  NOT  NULL , `flag1_ordine` tinyint( 1  )  DEFAULT NULL ,`ordinato` tinyint( 3  )  unsigned NOT  NULL DEFAULT  '0', `scalato_magazzino` tinyint( 3  )  unsigned NOT  NULL DEFAULT  '0',`richiedi_seriale` tinyint( 1  )  unsigned NOT  NULL DEFAULT  '0', `quantita_ordinato` decimal( 10, 3  )  NOT  NULL DEFAULT  '0.000', `quantita_impegnato` decimal( 10, 3  )  NOT  NULL DEFAULT  '0.000', `numeri_di_serie` int( 2  )  unsigned NOT  NULL DEFAULT  '0', `in_ordine` int( 3  )  NOT  NULL DEFAULT  '0', `pronto` int( 3  )  NOT  NULL DEFAULT  '0', `consegnato` int( 3  )  NOT  NULL DEFAULT  '0', PRIMARY  KEY (  `iddett`  ) , KEY  `idpro` (  `idpro`  ) , KEY  `idfat` (  `idfat`  ) ,KEY  `pro_e_var` (  `idpro` ,  `idvara` ,  `idvarb`  ) ,KEY  `ordine` (  `ordine`  )  ) ENGINE  =  MyISAM  DEFAULT CHARSET  = utf8;")

	call imposta_versione(nuova_versione)'395
end if

nuova_versione=396
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` ADD  `secondario` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	
	conn.execute ("DROP TABLE IF EXISTS password_reset")
	conn.execute ("CREATE TABLE IF NOT EXISTS `password_reset` ( `iduser` int(11) NOT NULL, PRIMARY KEY (`iduser`), CONSTRAINT FK_PersonOrder FOREIGN KEY (iduser) REFERENCES utenti(iduser)) ENGINE=MyISAM DEFAULT CHARSET=latin1;")
	
	conn.execute ("ALTER TABLE  `password_reset` ADD  `newpassword` varchar(255) NOT NULL default ''")
	
	call imposta_versione(nuova_versione)'396
end if
nuova_versione=397
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` ADD  `cognome` VARCHAR( 100 ) NULL DEFAULT '' after Data")
	conn.execute ("ALTER TABLE  `utenti` ADD  `nome` VARCHAR( 100 ) NULL DEFAULT '' after cognome")
	call imposta_versione(nuova_versione)'397
end if
nuova_versione=398
if versione<nuova_versione  then
	set rs_intestazioni=conn.execute("select i.cognome, i.nome, u.iduser from utenti u inner join utenti_intestazioni i on u.idintestazione = i.id")
	do while not rs_intestazioni.eof
	
		conn.execute("update utenti set cognome='"&duplica_apici(rs_intestazioni("cognome"))&"', nome='"&duplica_apici(rs_intestazioni("nome"))&"' where iduser="&rs_intestazioni("iduser"))
		response.write "aggiorno "&rs_intestazioni("cognome")&" iduser:"&rs_intestazioni("iduser")&"<br>"
		rs_intestazioni.movenext
	loop

	call imposta_versione(nuova_versione)'398
end if
nuova_versione=399
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` ADD  `ruolo` VARCHAR( 100 ) NULL DEFAULT '' after nome")
	call imposta_versione(nuova_versione)'399
end if
nuova_versione=400
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `fatture` ADD  `stato_fattura` INT( 1 ) UNSIGNED NOT NULL  DEFAULT '0' after iduser")
	call imposta_versione(nuova_versione)'400
end if
nuova_versione=401
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti_intestazioni` CHANGE  `azienda`  `azienda` VARCHAR( 200 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	conn.execute ("ALTER TABLE  `utenti_consegna` CHANGE  `d_azienda`  `d_azienda` VARCHAR( 200 ) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'401
end if

nuova_versione=402
if versione<nuova_versione  then
	
	conn.execute ("ALTER TABLE  `password_reset` ADD  ora_invio datetime NOT NULL")
	
	call imposta_versione(nuova_versione)'402
end if
nuova_versione=403
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` CHANGE  `secondario`  `secondario` INT( 11 ) UNSIGNED NOT NULL DEFAULT  '0'")
	
	
	set rs = conn.execute("select * from utenti where secondario>0")
	do while not rs.EOF
		response.write "modifico secondario:"&rs("iduser")&"<br>"
		iduser=conn.execute("select iduser from utenti_intestazioni where id ="&rs("idintestazione"))(0)
		conn.execute("update utenti set secondario="&iduser&" where iduser="&rs("iduser"))
		rs.movenext
	 Loop
	
	
	call imposta_versione(nuova_versione)'403
end if

nuova_versione=404
if versione<nuova_versione  then
	
	conn.execute ("ALTER TABLE  `ordini` ADD  note_cliente text default null")
	
	call imposta_versione(nuova_versione)'404
end if





nuova_versione=405
if versione<nuova_versione  then

	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `collant` VARCHAR( 10 ) NULL DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `polo` VARCHAR( 10 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'405
end if

nuova_versione=406
if versione<nuova_versione  then

	conn.execute ("ALTER TABLE  `scadenze` CHANGE  `importo`  `importo` DECIMAL( 10, 2 ) NULL DEFAULT NULL")
	call imposta_versione(nuova_versione)'406
end if

nuova_versione=407
if versione<nuova_versione  then

	conn.execute ("ALTER TABLE  `settori` ADD  `colore`  VARCHAR( 100 ) NULL DEFAULT ''")
	call imposta_versione(nuova_versione)'407
end if


nuova_versione=408
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_scarpa` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_collant` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_berretto` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_guanto` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_maglieria` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_polo` varchar(200) DEFAULT ''")
	conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_cintura` varchar(200) DEFAULT ''")
		conn.execute ("ALTER TABLE  `dipendenti_misure` ADD  `note_cinturone` varchar(200) DEFAULT ''")

	call imposta_versione(nuova_versione)'408
end if

nuova_versione=409
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `utenti` ADD  `fattura_sp` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini` ADD  `fattura_sp` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `fatture` ADD  `fattura_sp` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")

	call imposta_versione(nuova_versione)'409
end if
nuova_versione=410
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `preventivi` ADD  `fattura_sp` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")
	conn.execute ("ALTER TABLE  `ordini_fornitori` ADD  `fattura_sp` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")

	call imposta_versione(nuova_versione)'409
end if


nuova_versione=411
if versione<nuova_versione  then
	conn.execute ("ALTER TABLE  `fatture` ADD  `fattura_ce` TINYINT( 1 ) UNSIGNED NOT NULL DEFAULT '0'")

	call imposta_versione(nuova_versione)'411
end if







'				case "integer"			: sFieldType = "SMALLINT"
'				case "long"			: sFieldType = "INTEGER"
'				case "boolean"			: sFieldType = "BIT"
'				case "date"				: sFieldType = "DATE"
'				case "currency"			: sFieldType = "CURRENCY"
'				case "text"			: sFieldType = "VARCHAR"
'				case "memo"		: sFieldType = "TEXT"
'				case "ole"	: sFieldType = "LONGVARBINARY"
'				case "guid"				: sFieldType = "GUID"
'				case "byte"	: sFieldType = "TINYINT"
'				case "single"	: sFieldType = "SINGLE"

response.write "<br><b>VERSIONE FINALE: "&nuova_versione&"</b><br><br>"
%>