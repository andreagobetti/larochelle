<!--#include virtual="/setup.asp" -->
<!--#include virtual="/md5.asp" -->
<%

	  sql="select * from utenti order by iduser"
	  rs.Open sql, conn, 3, 3
	  rs.movefirst
	  do until rs.eof
	  if rs("password_md5")<> md5(rs("pass")) then
	  response.Write("<br>Errato: "&rs("iduser")&"<br>")
	  else
	  response.write "."
	  end if
	  rs.movenext
	  loop
	  rs.close
	  %>
