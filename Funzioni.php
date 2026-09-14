<?php
class MysqlClass
{
  // parametri per la connessione al database
  private $nomehost = "62.149.150.169";     
  private $nomeuser = "Sql594349";          
  private $password = "a125ee15"; 
  private $db = "Sql594349_2";
  public $conn;
          
  // controllo sulle connessioni attive
  private $attiva = false;
 
  // funzione per la connessione a MySQL
  public function connetti()
  {
   if(!$this->attiva)
   {
    //$connessione = mysql_connect($this->nomehost,$this->nomeuser,$this->password);
    $conn = mysql_connect($this->nomehost,$this->nomeuser,$this->password, $this->db) or die (mysql_error());
    
       }else{
        return true;
       }
    }
    
    // funzione per la chiusura della connessione
	public function disconnetti()
	{
	        if($this->attiva)
	        {
	                if(mysql_close())
	                {
	         $this->attiva = false; 
	             return true; 
	                }else{
	                        return false; 
	                }
	        }
	 }
	 
    //funzione per l'esecuzione delle query 
	public function query($sql)
	 {
	  if(isset($this->attiva))
	  {
	  $sql = mysql_query($sql) or die (mysql_error());
	  return $sql;
	  }else{
	  return false; 
	  }
	 }
}       
?>