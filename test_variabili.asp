<%
	
v=""
	
call a()


response.write "fine:"&v&"<br>"
	
	
sub a()
	v="a"
	response.write "a():"&v&"<br>"
	call b()
end sub	

sub b()
	v=v&"b"
	response.write "b():"&v&"<br>"
	
end sub

	
%>

