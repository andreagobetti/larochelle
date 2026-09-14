<%t_inizio=timer()%>
<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/chk_piva_cf.asp" -->
<!--#include virtual="/paginazione.asp" -->
<%
dim oper
dim main_page
main_page=false
if not utente_admin then call login()
idfat=request("idfat")
idord=request.form("idord")
nfat=request("nfat")
oper=lcase(request("oper"))
cercain=request("cercain")







'response.write oper
%>
<!--#include virtual="/sub_head_adm.asp" -->
<!--#include virtual="/regioni_inc.asp" -->



</head>
<body>
<div id="wrap">
<div id="header">
<%=titolo_top%> 
        <!-- Box CORPO INIZIO-->
		<%barra=6%>	<div id="barra_fissa">
			  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <%
	    if cercain="pa" then
			grigio="style='color:#999;'"
			document_title="Fatture PA"
		else
			grigiopa="style='color:#999;'"
			document_title="Fatture"
		end if
		%>

        <div class="ui-widget-header ui-corner-top titolo_admin">
			<a href="<%=questofile%>" <%=grigio%>>Fatture</a> / <a href="<%=questofile%>?cercain=pa" <%=grigiopa%>>Fatture PA</a>
	    </div>
	    </div>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
                <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>
        <p>
	        
	        
	        fsdfdsfdsfsd
        </p>


</div>
<%
	call connclose()
	txt_timer=txt_timer&"T fine:"&FormatNumber(timer() - t_inizio, 2)&" "
%>
<div id="footer"><%=txt_timer%></div>
<!--#include virtual="/pag_adm_footer_inc.asp" -->

</div>

</body>
</html>
