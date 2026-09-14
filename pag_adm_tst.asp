<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

if request.form("aggiorna")<>"" then
	sql="UPDATE Impostazioni SET "
	sql= sql & " Impostazioni.Top = '" & fixquotes(request.form("top")) & "'"
	sql= sql & ";"
	conn.Execute (Sql)
	Application("ConfigOk")= "ko"
end if
sql="select impostazioni.* FROM impostazioni "
rs=conn.execute(sql)

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<%
server.Execute(skins&"/stile.asp")
%>
<!--#include virtual="/sub_head.asp" -->
</head>
<body>
<center>
  <table height="100%" border="0" cellpadding="<%=cellp%>" cellspacing="<%=cells%>" class="tabellacentrale">     <tr align="center" valign="middle" > 
      <td height="30" colspan="3"> <!--#include virtual="/sub_top.asp" -->
 </td>
<tr> 
      <td width="<%=col_1%>" valign="top"> <!-- Colonna SINISTRA INIZIO-->
        <!--#include virtual="/sub_menu.asp" --><!-- Colonna SINISTRA FINE-->
      </td>
      <td width="<%=col_23%>" align="center" valign="top"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tabella1">
          <tr> 
            <td class=titolomenu>&nbsp;Amministra</td>
          </tr>
		  </table>
              <!-- Colonna CORPO INIZIO-->
              <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top=0px;">
                
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
              <td class="ui-widget-header"><a href="pag_adm_hp.asp">Home Page</a>&gt;Testata</td>
            </tr>
            <tr> 
              <td align="center">
<div align="left">&nbsp;<a href="javascript:editor('form1.top');"><b>Editor</b></a></div><br>
                  <textarea name="top" cols="60" rows="15" style="FONT: 12px;" ><%=rs("top")%></textarea>
              </td>
            </tr>
            <tr> 
              <td align="center"> <input type="submit" name="Aggiorna" value="Aggiorna"> 
              </td>
            </tr>
          </table>
              </form>
              <!-- Colonna CORPO FINE-->

      </td>
    </tr>
    <tr> 
      <td colspan="2" height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
