<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<%
if session("idadmin") = "" then call login()
anno=request.form("anno")
if anno="" then
	anno=year(date())
end if
set rs_o=conn.execute("select ordini.data FROM ordini LIMIT 0,1;")
min_year=year(rs_o("data"))
set rs_o=nothing
%>
<!--#include virtual="/sub_head_adm.asp" -->
    <link class="include" rel="stylesheet" type="text/css" href="jquery/css/jquery.jqplot.min.css" />
    <!--[if lt IE 9]><script language="javascript" type="text/javascript" src="jquery/js/excanvas.min.js"></script><![endif]-->
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%>
        <!-- Div Content INIZIO-->      <%barra=0%><div id="barra_fissa">
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
		<form id="form1" name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin:0px 0px 0px 0px;" >
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Bilancio</a>
		<select name="anno" onChange="this.form.submit()">
                <%
for n=min_year to year(date())
response.write "<option value='"&n&"'"
if n=cint(anno) then response.write " selected"
response.write  ">" & n&"</option>"
next
%>
              </select></div></form>    </div>  
              
              
              
<div class="ui-widget-header ui-corner-all" style="margin-top:10px; padding-top:5px; padding-bottom:5px;">Fatture clienti e fornitori</div>
        
                   
<%
'creo serieEntrate
serieEntrate=""
TotaleCol1=0
TotaleNoteCredito=0
Totaleentrate=0
dim col1(12)
dim noteCredito(12)
dim entrate(12)


for n=1 to 12
	Sql_ordini="select Sum(fatture.totale_merce) AS SommaDiTotale FROM fatture WHERE (Month(data)="&n&") AND (anno="&anno&") ;	"
	set rs=conn.execute(Sql_ordini)
	if not isnull(rs("SommaDiTotale")) then
		fatture=cdbl(rs("SommaDiTotale"))
		
		TotaleCol1=TotaleCol1+fatture
		
	else
		fatture=0
	end if
	
	col1(n)=fatture
	
	
	
	'Note di credito
	Sql="select Sum(ordini.totale_merce_ordine) AS totale_merce_ordine, Sum(ordini.trasporto) as totaleTrasporto FROM ordini WHERE (Month(data)="&n&") AND (year(data)="&anno&") and tipo_documento='notacredito' and eliminato=0"	set rs=conn.execute(Sql)
	if not isnull(rs("totale_merce_ordine")) then
		SommaDiTotale=cdbl(rs("totale_merce_ordine"))
		TotaleNoteCredito=TotaleNoteCredito+SommaDiTotale
	else
		SommaDiTotale="0"
	end if
	noteCredito(n)=SommaDiTotale

	entrate(n)=fatture-SommaDiTotale
	
	
	TotaleEntrate=TotaleEntrate+entrate(n)
	
	serieEntrate=serieEntrate&FormatNumber(entrate(n),0,0,0,0)
	if n<12 then serieEntrate=serieEntrate&"," 'aggiungo la virgola

	
	
	
	
	
	
next


'creo serieFattureFornitori
dim col2(12)

serieFattureFornitori=""
TotaleCol2=0
for n=1 to 12
	Sql_scadenze_for="select Sum(fatture_for.Totale_fattura) AS SommaDiimporto FROM fatture_for WHERE ((Month(data)="&n&") AND (year(data)="&anno&"));"
	set rs=conn.execute(Sql_scadenze_for)
	if not isnull(rs("SommaDiimporto")) then
		SommaDiimporto=rs("SommaDiimporto")
		'Scorporo iva
		SommaDiimporto=cdbl(SommaDiimporto)/1.22
		TotaleCol2=TotaleCol2+SommaDiimporto
		
	else
		SommaDiimporto=0
	end if
	
	col2(n)=SommaDiimporto

	serieFattureFornitori=serieFattureFornitori&FormatNumber(SommaDiimporto,0,0,0,0)

	if n<12 then serieFattureFornitori=serieFattureFornitori&"," 'aggiungo la virgola
next

'creo s3
s3=""
dim col3(12)
TotaleCol3=0
for n=1 to 12


	'Totale ordini nel mese
	Sql_ordini="select Sum(ordini.totale_merce_ordine) AS totale_merce_ordine, Sum(ordini.trasporto) as totaleTrasporto FROM ordini WHERE (Month(data)="&n&") AND (year(data)="&anno&") and tipo_documento='ordine' and eliminato=0"	set rs=conn.execute(Sql_ordini)
	if not isnull(rs("totale_merce_ordine")) then
		totale_imponibile_ordine=cdbl(rs("totale_merce_ordine"))+cdbl(rs("totaleTrasporto"))
		
		TotaleCol3=TotaleCol3+totale_imponibile_ordine
	else
		totale_imponibile_ordine=0
	end if
	
	s3=s3&FormatNumber(totale_imponibile_ordine,0,0,0,0)

	
	
	
	
	
	col3(n)=totale_imponibile_ordine

	if n<12 then s3=s3&"," 'aggiungo la virgola
next






Set rs = Nothing
'serieFattureFornitori=serieFattureFornitori&","&int(somma_tot/10)
%>


    <div id="chart1" style="margin-top:10px; margin-bottom:30px; margin-left:20px; width:95%; height:400px;"></div>
	
	<p>Entrate: totale imponibile fatture clienti - totale imponibile note credito con data emissione nel mese</p>
	<p>Ordini clienti: totale imponibile ordini clienti con data emissione nel mese
    <p>Fatture fornitori: totale imponibile fatture da fornitori con data emissione nel mese</p>

   <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	<tr>
		<td>Mese
		</td>
		<td align="right">Fatture clienti:
		</td>
		<td align="right">Note credito:
		</td>
		<td align="right">Entrate:
		</td>
		
		<td align="right">Ordini
		</td>
		<td align="right">fatture fornitori
		</td>
		</tr>
	<%
		for n=1 to 12
	 %>
	<tr>
		<td><%=nome_mese(n)%>
		</td>
		<td align="right"><a href="pag_adm_rep07.asp?mese=<%=n%>"><%=formatnumber(col1(n),2)%></a>
		</td>
		<td align="right"><%=formatnumber(notecredito(n),2)%>
		</td>
		<td align="right"><%=formatnumber(entrate(n),2)%>
		</td>
		<td align="right"><a href="pag_adm_rep10.asp?mese=<%=n%>"><%=formatnumber(col3(n),2)%></a>
		</td>		
		
		<td align="right"><a href="pag_adm_rep18.asp?mese=<%=n%>"><%=formatnumber(col2(n),2)%></a>
		</td>
		</tr>
	<%next %>
	<tr >
		<td><strong>Totale</strong>
		</td>
		<td align="right"><strong><%=formatnumber(TotaleCol1,2)%></strong>
		</td>
		<td align="right"><strong><%=formatnumber(Totalenotecredito,2)%></strong>
		</td>
		<td align="right"><strong><%=formatnumber(TotaleEntrate,2)%></strong>
		</td>
		
		<td align="right"><strong><%=formatnumber(TotaleCol3,2)%></strong>
		</td>
		
		<td align="right"><strong><%=formatnumber(TotaleCol2,2)%></strong>
		</td>
		</tr>
	</table>
	
      <script class="code" type="text/javascript">
$(document).ready(function(){
    var serieEntrate = [<%=serieEntrate%>];
    var s3 =[<%=s3%>];
    var serieFattureFornitori = [<%=serieFattureFornitori%>];
    // Can specify a custom tick Array.
    // Ticks should match up one for each y value (category) in the series.
    var ticks = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
     
    var plot1 = $.jqplot('chart1', [serieEntrate,serieFattureFornitori], {
        // The "seriesDefaults" option is an options object that will
        // be applied to all series in the chart.
        seriesDefaults:{
            renderer:$.jqplot.BarRenderer,
            rendererOptions: {fillToZero: true}
        },
        // Custom labels for the series are specified with the "label"
        // option on the series option.  Here a series option object
        // is specified for each series.
        series:[
            {label:'Entrate'},
            {label:'Fatture fornitori'},{
	            label:'ordini'
            }
        ],
        // Show the legend and put it outside the grid, but inside the
        // plot container, shrinking the grid to accomodate the legend.
        // A value of "outside" would not shrink the grid and allow
        // the legend to overflow the container.
        legend: {
            show: true,
            //placement: 'outsideGrid'
			placement: 'insideGrid'
        },
        axes: {
            // Use a category axis on the x axis and use our custom ticks.
            xaxis: {
                renderer: $.jqplot.CategoryAxisRenderer,
                ticks: ticks
            },
            // Pad the y axis just a little so bars can get close to, but
            // not touch, the grid boundaries.  1.2 is the default padding.
            yaxis: {
                pad: 1.05,
                tickOptions: {formatString: '%d'}
            }
        }
    });
});
      </script>
      
      
      <!-- SECONDO GRAFICO -->
      
      
      
      
      
      
      
      
      

      
      
      
      <script class="include" type="text/javascript" src="jquery/js/jquery.jqplot.min.js"></script>
      <script type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.barRenderer.min.js"></script>
      <script type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.categoryAxisRenderer.min.js"></script>
      <script type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.pointLabels.min.js"></script>
      <script class="include" type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.pieRenderer.min.js"></script>
</div>
<div id="footer"></div>
<%
	
	call connclose()
	%>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>