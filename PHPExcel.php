<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';



$iduser=$_GET['iduser'];
include ("/config/mdb.php");


// Recupero i dati azienda dal db
$sql="select u.iduser, i.azienda from utenti u left join utenti_intestazioni i on u.idintestazione = i.id where u.iduser={$iduser}";

if (!$result = $conn->query($sql)) {
  echo "Errore della query: " . $conn->error . ".";
  exit();
}else{
  // conteggio dei record
  if($result->num_rows > 0) {
   $row = $result->fetch_array(MYSQLI_ASSOC);
  // liberazione delle risorse occupate dal risultato
    $result->close();
  }
}

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Ordine');

// Set document properties
$objPHPExcel->getProperties()->setCreator("La Rochelle")
							 ->setLastModifiedBy("La Rochelle")
							 ->setTitle("Modulo ordine precompilato")
							 ->setSubject("Cliente ".$row['azienda'])
							 ->setDescription("Modulo ordine precompilato ".$row['azienda']);
							// ->setKeywords("office 2007 openxml php")
							// ->setCategory("Test result file");
							
							
							
							
							
							

//Immagine
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Logo');
$objDrawing->setDescription('Logo');
$objDrawing->setPath('logo_documenti.jpg');
//$objDrawing->setResizeProportional(true);
$objDrawing->setWidthAndHeight(148,100);

$objDrawing->setCoordinates('A1');
//$objDrawing->setOffsetX(110);
//$objDrawing->setRotation(25);
//$objDrawing->getShadow()->setVisible(true);
//$objDrawing->getShadow()->setDirection(45);
$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());


$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);

//Intestazioni
$row=5;



$objPHPExcel->getActiveSheet()
            ->setCellValue('A'.$row, 'Codice')
            ->setCellValue('B'.$row, 'Articolo')
            ->setCellValue('C'.$row, 'Prezzo');
            
$styleArray = array(
'font' => array(
'bold' => true,'size'=>16
)
);

$objPHPExcel->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B'.$row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('C'.$row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('C'.$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);




//Recupero elenco dipendenti
$sql="select dipendenti.*, misure.data_rilievo, utenti_gradi.nome_grado, utenti_gradi.colore, utenti_gradi.colore_nominativo FROM (dipendenti left JOIN utenti_gradi ON dipendenti.grado = utenti_gradi.id) left join (SELECT iddip,max(data_rilievo) as data_rilievo FROM dipendenti_misure group by iddip  ) as misure ON dipendenti.iddip = misure.iddip inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.iduser={$iduser}";

$col=3; //Colonna iniziale elenco dipendenti
if (!$dipendenti = $conn->query($sql)) {
  echo "Errore della query: " . $conn->error . ".";
  exit();
}else{
  // conteggio dei record
  $num_dipendenti=$dipendenti->num_rows;
  if($num_dipendenti > 0) {
    // conteggio dei record restituiti dalla query
    while($dipendente = $dipendenti->fetch_array(MYSQLI_ASSOC))
    {
	    
	    // Add some data
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $dipendente['cognome']." ".$dipendente['nome']);
	    
	    //Ruota testo 90°
	    $objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->getAlignment()->setTextRotation(90);
	    //Allineamento orizzontale centrato
	    $objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

	    
	    
	    $keyCell = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col,$row);
	    
	    
	    
		//$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row)->getAlignment()->setTextRotation(90);
	//	$objPHPExcel->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
	
	
	
	
	    $col++;
	    
    }
    // liberazione delle risorse occupate dal risultato
    $dipendenti->close();
  }
  
}

//Aggiungo quantita totale
	    // Add some data
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Quantità totale');
	    
	    //Ruota testo 90°
	    $objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->getAlignment()->setTextRotation(90);
	    //Allineamento orizzontale centrato
	    $objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->applyFromArray($styleArray);
//Aggiungo importo totale
		$col++;
	    // Add some data
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Totale');
	    
	    //Ruota testo 90°
	    //$objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->getAlignment()->setTextRotation(90);
	    //Allineamento orizzontale centrato
	    //$objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle(getColFromNumber($col).$row)->applyFromArray($styleArray);








//Sfondo prima riga tabella
$objPHPExcel->getActiveSheet()
		    ->getStyle('A'.$row.':'.getColFromNumber($col).$row)
		    ->getFill()
		    ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
		    ->getStartColor()
		    ->setARGB('FF808080');



//Recupero articoli nel carrello
$sql="select carrello.*, prodotti.codice, prodotti.articolo, prodotti.variante1, prodotti.variante2, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata, varianti_a.variante_a, varianti_a.prezzo_ve_va, varianti_b.variante_b, prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FROM varianti_b RIGHT JOIN (varianti_a RIGHT JOIN (carrello INNER JOIN prodotti ON carrello.idpro = prodotti.IDpro) ON varianti_a.IDvara = carrello.idvara) ON varianti_b.IDvarb = carrello.idvarb where iduser={$iduser}";

$col=1; //Colonna iniziale elenco dipendenti
$row++;
if (!$articoli = $conn->query($sql)) {
  echo "Errore della query: " . $conn->error . ".";
  exit();
}else{
  // conteggio dei record
  if($articoli->num_rows > 0) {
    // conteggio dei record restituiti dalla query
    while($articolo = $articoli->fetch_array(MYSQLI_ASSOC))
    {
	    
	    // Add some data
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, $articolo['codice']);
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1, $row, $articolo['articolo']);
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, $articolo['prezzo']);
	    
	    //Aggiungo formula per totale quantità dipendenti
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3+$num_dipendenti, $row, '=SUM(D'.$row.':'.getColFromNumber(2+$num_dipendenti).$row.')');
		$objPHPExcel->getActiveSheet()->getStyle(getColFromNumber(3+$num_dipendenti).$row)->applyFromArray(array('font' => array('bold' => true,'size'=>13)));
	    //Aggiungo formula per totale prezzo
	    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4+$num_dipendenti, $row, '='.getColFromNumber(3+$num_dipendenti).$row.'*C'.$row.')');
		$objPHPExcel->getActiveSheet()->getStyle(getColFromNumber(4+$num_dipendenti).$row)->applyFromArray(array('font' => array('bold' => true,'size'=>13)));
	    
	    
	    
	    
	
	
	    $row++;
	    
    }
    // liberazione delle risorse occupate dal risultato
    $articoli->close();
  }
}







// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;

function addressbyrowcol($row,$col) {
return getColFromNumber($col).$row;
}


function getColFromNumber($num) {
    $numeric = ($num ) % 26;
    $letter = chr(65 + $numeric);
    $num2 = intval(($num - 1) / 26);
    if ($num2 > 0) {
        return getColFromNumber($num2) . $letter;
    } else {
        return $letter;
    }
}

