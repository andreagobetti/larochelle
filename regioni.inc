<%
function regione(n)
select case n
case -1
regione=22
case 0
regione=""
case 1
regione="Piemonte"
case 2
regione="Valle D'Aosta"
case 3
regione="Lombardia"
case 4
regione="Trentino Alto Adige"
case 5
regione="Veneto"
case 6
regione="Friuli Venezia Giulia"
case 7
regione="Liguria"
case 8
regione="Emilia Romagna"
case 9
regione="Toscana"
case 10
regione="Umbria"
case 11
regione="Marche"
case 12
regione="Lazio"
case 13
regione="Abruzzo"
case 14
regione="Molise"
case 15
regione="Campania"
case 16
regione="Puglia"
case 17
regione="Basilicata"
case 18
regione="Calabria"
case 19
regione="Sicilia"
case 20
regione="Sardegna"
case 21
regione="Stato estero"
case 22
regione="Repubblica di San Marino"
end select
end function
%>