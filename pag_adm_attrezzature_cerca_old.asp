<!--#include virtual="/setup.asp" -->
<script type="text/javascript">


$arraysquadre=array();

########	ESEMPIO CON ARRAY	##################

<%
sql="select * from utenti order by user "
set rs=conn.execute(sql)
do until rs.EOF
%>array_push($arraysquadre,'<%=rs("user")%>');
<%
	rs.movenext
loop
%>



sort($arraysquadre);
</script>

<ul>
<script type="text/javascript">
for($a=0;$a<count($arraysquadre);$a++){
	if(substr_count($arraysquadre[$a],$squadra)>0){
		echo "<li>".$arraysquadre[$a]."</li>";
	}

}
</script>
</ul>
