<?php
header('Content-type: application/json');	
$idcategoria=$_GET['idcategoria'];	
$filename=$_GET['filename'];

//$filename="borse_.csv";
//$idcategoria=11;



$array_csv=(csv_to_array($_SERVER['DOCUMENT_ROOT'] .'/public/mepa_articoli/'.$filename, ';'));

//Connessione al db

include ("/config/mdb.php");
include ("functions.php");


//$conn = new mysqli('62.149.150.169', 'Sql594349', 'a125ee15', 'Sql594349_2');
if ($conn->connect_error) {
    die('Errore di connessione (' . $mysqli->connect_errno . ') '
    . $mysqli->connect_error);
} 
//$conn->query("delete from prodotti");
//$conn->query("delete from prodotti_mepa_campi");

//Carico l'elenco dei campi
$result = $conn->query("select label,id from mepa_campi where idcategoria={$idcategoria}");
$campi=array();
while($rowcampi = $result->fetch_row()) {
  $campi[$rowcampi[0]]=$rowcampi[1];
}
//print_r($campi);


//exit;
$n=0;
$txt_log="";
foreach($array_csv as $row){
	$n++;
	
	$prezzo = floatval(str_replace(',', '.', str_replace('.', '', $row['Prezzo'])));
	$articolo = $conn->real_escape_string( $row['Nome commerciale'] . " " . $row['Descrizione tecnica']);
	
	
	$sql="INSERT INTO prodotti (data, datamod, codice, articolo, prezzo, um) VALUES (now(),now(),'{$row['Codice articolo fornitore']}','{$articolo}', {$prezzo},'{$row['Unità di misura']}')";
	
	// Esecuzione della query e controllo degli eventuali errori
	if (!$conn->query($sql)) {
		
		$result = $conn->query("select prodotti.idpro as idpro from prodotti where codice='{$row['Codice articolo fornitore']}'");
		$data = $result->fetch_assoc();
		$idpro['idpro'];
		$txt_log= $txt_log.$n.")".$row['Codice articolo fornitore']." già esistente<br>";
	    
	    
	}else{
		//print("Query eseguita");
		//echo "Righe generate: " .  $conn->affected_rows . "<br />";
		$idpro=$conn->insert_id ;
		$txt_log= $txt_log.$n.")".$row['Codice articolo fornitore']." aggiunto<br>";
		
		//echo "Ultimo ID inserito: " . $idpro  . "<br />";
	}
	
	//Converto i nomi dei campi in id
	$newrow=array();
	
	foreach ($row as $key => $value){
		$id=$campi[$key];
		//echo($key."=".$id."<br>");
		$newrow[$id]=utf8_encode($value);
	}	
	
	
	$json = $conn->real_escape_string( json_encode($newrow));
	
	$sql="INSERT INTO prodotti_mepa_campi (idcategoria, idpro, mepa_campi) VALUES ({$idcategoria},{$idpro},'{$json}')";
	// Esecuzione della query e controllo degli eventuali errori
	if (!$conn->query($sql)) {
		//echo $sql."<br>";
	    //die($conn->error);
	}else{
		//
	}
}
//add2log($txt_log, 0);
		$txt_log = $conn->real_escape_string( $txt_log);
		$sql="insert into log (data, evento, causale) values (now(), '{$txt_log}', 0)";
		$conn->query($sql);


$array=array("files"=>array("name"=>"","size"=>"","url"=>"","thumbnail_url"=>"","delete_url"=>"","delete_type"=>"POST","idfile"=>""));
echo json_encode($array);




	
function csv_to_array($filename='', $delimiter=';')
{
    if(!file_exists($filename) || !is_readable($filename))
        return FALSE;

    $header = NULL;
    $data = array();
    if (($handle = fopen($filename, 'r')) !== FALSE)
    {
        while (($row = fgetcsv($handle, 1000, $delimiter)) !== FALSE)
        {
            if(!$header){
                $header = $row;

                }
            else
                $data[] = array_combine($header, $row);
        }
        fclose($handle);
    }
    return $data;
}

?>