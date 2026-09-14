<%

a=null
response.write "NULL:"&isnull(a)&"<br>"

response.write "NULL:"&isnumeric(a)&"<br>"
response.write ":"&isnumeric("")&"<br>"
response.write "a:"&isnumeric("a")&"<br>"
response.write "5:"&isnumeric("5")&"<br>"



%>