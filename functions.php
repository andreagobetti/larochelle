<?php
	function add2log($testo, $gravita){
		//$testo = $conn->real_escape_string( $testo);
		$sql="insert into log (data, evento, causale) values (now(), '{$testo}', {$gravita})";
		$conn->query($sql);
	}
?>