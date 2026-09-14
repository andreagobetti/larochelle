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
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Grafici entrate e uscite</a>
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
'creo s1
s1=""
somma_tot=0
dim col1(12)
dim col2(12)
for n=1 to 12
	Sql_ordini="select Sum(fatture.Totale_fattura) AS SommaDiTotale FROM fatture WHERE (Month(data)="&n&") AND (anno="&anno&") ;	"
	set rs=conn.execute(Sql_ordini)
	if not isnull(rs("SommaDiTotale")) then
		somma_tot=somma_tot+cdbl(rs("SommaDiTotale"))
		SommaDiTotale=FormatNumber(cdbl(rs("SommaDiTotale")),0,0,0,0)
	else
		SommaDiTotale="0"
	end if
	col1(n)=SommaDiTotale
	s1=s1&SommaDiTotale
	if n<12 then s1=s1&"," 'aggiungo la virgola
next
's1=s1&","&int(somma_tot/10)
'creo s2
s2=""
somma_tot=0
for n=1 to 12
	Sql_scadenze_for="select Sum(fatture_for.Totale_fattura) AS SommaDiimporto FROM fatture_for WHERE ((Month(data)="&n&") AND (year(data)="&anno&"));"
	set rs=conn.execute(Sql_scadenze_for)
	if not isnull(rs("SommaDiimporto")) then
		SommaDiimporto=rs("SommaDiimporto")
		s2=s2&FormatNumber(cdbl(rs("SommaDiimporto")),0,0,0,0)
	else
		s2=s2&"0"
		SommaDiimporto=0
	end if
	col2(n)=SommaDiimporto

	if not isnull(rs("SommaDiimporto")) then somma_tot=somma_tot+clng(rs("SommaDiimporto"))
	if n<12 then s2=s2&"," 'aggiungo la virgola
next
Set rs = Nothing
's2=s2&","&int(somma_tot/10)
%>

	<p>Entrate: totale fatture clienti con data emissione nel mese</p>
    <p>Uscite: totale fatture da fornitori con data emissione nel mese</p>

    <div id="chart1" style="margin-top:10px; margin-bottom:30px; margin-left:20px; width:95%; height:400px;"></div>
	
	
               <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	<tr>
		<td>Mese
		</td>
		<td align="right">Entrate
		</td>
		<td align="right">Uscite
		</td>
		</tr>
	<%
		tot_col1=0
		tot_col2=0
		for n=1 to 12
		tot_col1=tot_col1+cdbl(col1(n))
		tot_col2=tot_col2+cdbl(col2(n))
	 %>
	<tr>
		<td><%=nome_mese(n)%>
		</td>
		<td align="right"><a href="pag_adm_rep07.asp?mese=<%=n%>"><%=formatnumber(col1(n),2)%></a>
		</td>
		<td align="right"><a href="pag_adm_rep18.asp?mese=<%=n%>"><%=formatnumber(col2(n),2)%></a>
		</td>
		</tr>
	<%next %>
	<tr >
		<td><strong>Totale</strong>
		</td>
		<td align="right"><strong><%=formatnumber(tot_col1,2)%></strong>
		</td>
		<td align="right"><strong><%=formatnumber(tot_col2,2)%></strong>
		</td>
		</tr>
	</table>
	
      <script class="code" type="text/javascript">
$(document).ready(function(){
    var s1 = [<%=s1%>];
    var s2 = [<%=s2%>];
    // Can specify a custom tick Array.
    // Ticks should match up one for each y value (category) in the series.
    var ticks = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
     
    var plot1 = $.jqplot('chart1', [s1, s2], {
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
            {label:'Uscite'}
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
      
<div class="ui-widget-header ui-corner-all" style="margin-top:10px; padding-top:5px; padding-bottom:5px;">Incassi fatture clienti e scadenze fatture fornitori</div>

      <%
'creo s1
s1=""
somma_tot=0
for n=1 to 12
	'Totale incassi
	Sql_ordini="select Sum(incassi.importo) AS SommaDiTotale FROM incassi WHERE (Month(data)="&n&") AND (year(data)="&anno&") 	"	set rs=conn.execute(Sql_ordini)
	if not isnull(rs("SommaDiTotale")) then
		somma_tot=somma_tot+cdbl(rs("SommaDiTotale"))
		SommaDiTotale=cdbl(rs("SommaDiTotale"))
	else
		SommaDiTotale=0
	end if
	'Totale scadenze
	Sql_ordini="select Sum(scadenze.importo) AS SommaDiTotale FROM scadenze WHERE (Month(data)="&n&") AND (year(data)="&anno&") 	"	set rs=conn.execute(Sql_ordini)
	if not isnull(rs("SommaDiTotale")) then
		somma_tot=somma_tot+cdbl(rs("SommaDiTotale"))
		SommaDiTotale=SommaDiTotale+cdbl(rs("SommaDiTotale"))
	else
		'SommaDiTotale="0"
	end if
	
	

	
	
	
	col1(n)=SommaDiTotale
	SommaDiTotale=FormatNumber(SommaDiTotale,0,0,0,0)
	
	s1=s1&SommaDiTotale
	if n<12 then s1=s1&"," 'aggiungo la virgola
next
's1=s1&","&int(somma_tot/10)
'creo s2
s2=""
somma_tot=0
for n=1 to 12
	SommaDiimporto=0
	Sql_scadenze_for="select Sum(scadenze_for.importo) AS SommaDiimporto FROM scadenze_for inner join fatture_for on scadenze_for.idfat = fatture_for.idfat WHERE idfor>0 and ((Month(scadenze_for.scadenza)="&n&") AND (year(scadenze_for.scadenza)="&anno&"));"
	set rs=conn.execute(Sql_scadenze_for)
	if not isnull(rs("SommaDiimporto")) then
		SommaDiimporto=rs("SommaDiimporto")
		s2=s2&FormatNumber(cdbl(rs("SommaDiimporto")),0,0,0,0)
	else
		s2=s2&"0"
	end if
	col2(n)=SommaDiimporto

	if not isnull(rs("SommaDiimporto")) then somma_tot=somma_tot+clng(rs("SommaDiimporto"))
	if n<12 then s2=s2&"," 'aggiungo la virgola
next
Set rs = Nothing
's2=s2&","&int(somma_tot/10)
%>

	<p>Entrate: totale incassi e scadenze fatture clienti nel mese</p>
    <p>Uscite: totale scadenze fatture da fornitori nel mese</p>

    <div id="chart2" style="margin-top:10px; margin-bottom:30px; margin-left:20px; width:95%; height:400px;"></div>
	
	
               <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	<tr>
		<td>Mese
		</td>
		<td align="right">Entrate
		</td>
		<td align="right">Uscite
		</td>
		</tr>
	<%
		tot_col1=0
		tot_col2=0
		for n=1 to 12
		tot_col1=tot_col1+cdbl(col1(n))
		tot_col2=tot_col2+cdbl(col2(n))
	 %>
	<tr>
		<td><%=nome_mese(n)%>
		</td>
		<td align="right"><%=formatnumber(col1(n),2)%>
		</td>
		<td align="right"><a href="pag_adm_rep15.asp?tutte=si&mese=<%=n%>"><%=formatnumber(col2(n),2)%></a>
		</td>
		</tr>
	<%next %>
	<tr >
		<td><strong>Totale</strong>
		</td>
		<td align="right"><strong><%=formatnumber(tot_col1,2)%></strong>
		</td>
		<td align="right"><strong><%=formatnumber(tot_col2,2)%></strong>
		</td>
		</tr>
	</table>
	<%
		call connclose()

%>
      <script class="code" type="text/javascript">
$(document).ready(function(){
    var s1 = [<%=s1%>];
    var s2 = [<%=s2%>];
    // Can specify a custom tick Array.
    // Ticks should match up one for each y value (category) in the series.
    var ticks = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
     
    var plot1 = $.jqplot('chart2', [s1, s2], {
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
            {label:'Uscite'}
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

      
      
      
      <script class="include" type="text/javascript" src="jquery/js/jquery.jqplot.min.js"></script>
      <script type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.barRenderer.min.js"></script>
      <script type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.categoryAxisRenderer.min.js"></script>
      <script type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.pointLabels.min.js"></script>
      <script class="include" type="text/javascript" src="jquery/js/jqplot.plugins/jqplot.pieRenderer.min.js"></script>
</div>
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>