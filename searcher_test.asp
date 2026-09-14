<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<%    
Response.ContentType = "application/json; charset=utf-8"
'Response.ContentType = "text/html"
'Response.AddHeader "Content-Type", "application/json;charset=UTF-8"
Response.CodePage = 65001
'Response.CharSet = "UTF-8"

For Each item In Request.Querystring
    str_add2log= str_add2log&"(" & item & "):" & Request.Querystring(item) & "|"
Next
session("str_add2log")=str_add2log




%>
