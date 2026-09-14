<?php



include ("/config/mdb.php");
/*
  selezione di dati da una tabella con MySQLi
*/
// inclusione del file di connessione

// esecuzione della query per la selezione dei record
// query argomento del metodo query()
$sql="select u.iduser, i.azienda from utenti u left join utenti_intestazioni i on u.idintestazione = i.id";

if (!$result = $conn->query($sql)) {
  echo "Errore della query: " . $conn->error . ".";
  exit();
}else{
  // conteggio dei record
  if($result->num_rows > 0) {
    // conteggio dei record restituiti dalla query
    while($row = $result->fetch_array(MYSQLI_ASSOC))
    {
      echo $row['iduser']." ".$row['azienda']  . "<br />";
    }
    // liberazione delle risorse occupate dal risultato
    $result->close();
  }
}
// chiusura della connessione
$conn->close();

?>