<%

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=mid(Request.Servervariables("HTTP_HOST"),5)%>>Gestione portale</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<link href="admin_stile.css" rel="stylesheet" type="text/css">
</head>
<body>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td class=titadmin> <strong>Informazioni 
      server </strong> </td>
  </tr>
    <td align="center" valign="top"> 
      <table border="1" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Variable&nbsp;Name</b></font></td>
          <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Value</b></font></td>
        </tr>
        <% for each key in Request.ServerVariables %>
        <tr> 
          <td bgColor="<% =strPopUpTableColor %>" valign="top"><font size="2>"><b> 
            <% =key %>
            </b></font></td>
          <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="2"> 
            <%if Request.ServerVariables(key) = "" then
Response.Write "&nbsp;"
else
Response.Write Request.Servervariables(key)
end if 
%>
            </font></td>
        </tr>
        <% next %>
      </table>
</table>
</body>
</html>
