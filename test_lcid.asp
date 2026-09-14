<%
response.write("<p>")
response.write("Default LCID is: " & Session.LCID & "<br>")
response.write("Date format is: " & date() & "<br>")
response.write("Currency format is: " & FormatCurrency(350))
response.write("</p>")

Session.LCID=1036

response.write("<p>")
response.write("LCID is now: " & Session.LCID & "<br>")
response.write("Date format is: " & date() & "<br>")
response.write("Currency format is: " & FormatCurrency(350))
response.write("</p>")

Session.LCID=3079

response.write("<p>")
response.write("LCID is now: " & Session.LCID & "<br>")
response.write("Date format is: " & date() & "<br>")
response.write("Currency format is: " & FormatCurrency(350))
response.write("</p>")

Session.LCID=2057

response.write("<p>")
response.write("LCID is now: " & Session.LCID & "<br>")
response.write("Date format is: " & date() & "<br>")
response.write("Currency format is: " & FormatCurrency(350))&"<br>"

response.write FormatNumber(10,2)&"<br>"
response.write("</p>")
%>