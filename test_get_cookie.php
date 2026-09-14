
<?php


$tmp  = explode( '&', $_COOKIE['gestionale_larochelle_it']);
$num=count($tmp);
$data = array(array());
foreach ( $tmp as $k => $v )
{
	$tmp2=	explode( '=', $v );
  $data[$tmp2[0]] =$tmp2[1] ;
}
echo("data:".count($data)."<br>");
echo($data['iduser']);
?>
		
