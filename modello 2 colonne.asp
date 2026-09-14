<!--#include virtual="/setup.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<!-- Stile: <%=Application("skins")%> -->

<!--#include virtual="/sub_head.asp" -->
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%> </div>
<div id="sidebar"><!--#include virtual="/sub_menu.asp" --> </div>
<div id="content_large">
        <!-- Div Content INIZIO-->
        <div class="Aministra"><a href="pag_adm_main.asp" class="Aministra_rosso" ><span class="rosso">Amministrazione</span></a>&gt;<a href="<%=questofile%>" class="Amministra_piccolo">Home page</a></div>
        <!-- Div Content FINE-->
</div>
<div id="footer"></div><!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%
rsClose
%>
