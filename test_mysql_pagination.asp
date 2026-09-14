<%
Dim conn, rs
	strConn = "server=62.149.150.165;uid=Sql582810;pwd=e6bebde5;database=Sql582810_2;driver=MySQL ODBC 3.51 Driver"
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	conn.Open strConn
    Set rs = Server.CreateObject("ADODB.Recordset")

    Dim pag
    pag = Request.Querystring("pag")
    If IsNumeric(pag) = False Or pag < 1 Then pag = 1

    rs.Open "select COUNT(*) FROM log", conn, 1
    Dim quanti, fine, inizio
    quanti = clng(rs(0))
    rs.Close

    fine = 10
    inizio = (pag - 1) * fine

    rs.Open "select SQL_CALC_FOUND_ROWS * FROM log LIMIT " & inizio & ", " & fine, conn, 1
    response.write "select SQL_CALC_FOUND_ROWS * FROM log LIMIT " & inizio & ", " & fine
    lngTotalRecords = conn.Execute("Select Found_Rows();")(0).Value
    response.write "lngTotalRecords: "&lngTotalRecords
    While rs.EOF = False
        Response.Write "<p>" & rs("idlog") & "</p>"
        rs.MoveNext
    Wend
    rs.Close

    Call connClose()

    Dim i, k, max_left_right, intero
    max_left_right = 3
    response.write "quanti:"&quanti&" fine:"&fine
    intero = int(quanti / fine)
    
    If quanti > (intero * fine) Then
        intero = intero + 1
    End If
    If CInt(pag) > max_left_right Then
        i = pag - max_left_right
    Else
        i = 1
    End If
    If CInt(pag + max_left_right) < intero Then
        k = pag + max_left_right
    Else
        k = intero
    End If
%>
<p>Pagina <%=pag%> di <%=intero%></p>
<%
    If intero > 1 Then
%>
<p>
<%
        If i > 1 Then
%>
... |
<%
        End If
        Do While Not i > k
            If CInt(i) = CInt(pag) Then
%>
<b><%=i%></b>
<%
            Else
                If CInt(i) = 1 Then
%>
<a href="paginazione.asp"><%=i%></a>
<%
                Else
%>
<a href="paginazione.asp?pag=<%=i%>"><%=i%></a>
<%
                End If
            End If
            If i < intero Then
%>
|
<%
            End If
            i = i + 1
        Loop
        If k < intero Then
%>
...
<%
        End If
%>
</p>
<%
    End If
sub connclose()
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
end sub
%>