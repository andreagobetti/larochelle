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
date_default_timezone_set('Europe/Rome');



if (PHP_SAPI == 'cli')
    die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

$tmp = explode('&', $_COOKIE['gestionale_larochelle_it']);
$num = count($tmp);
$data = array(array());
foreach ($tmp as $k => $v) {
    $tmp2 = explode('=', $v);
    $data[$tmp2[0]] = $tmp2[1];
}

$iduser = $_GET['iduser'];
include ("/config/mdb.php");


// Recupero i dati azienda dal db
$sql = "select u.iduser, i.azienda, u.vedi_prezzi from utenti u left join utenti_intestazioni i on u.idintestazione = i.id where u.iduser={$iduser}";

if (!$result = $conn->query($sql)) {
    echo "Errore della query: " . $conn->error . ".";
    exit();
} else {
    // conteggio dei record
    if ($result->num_rows > 0) {
        $row = $result->fetch_array(MYSQLI_ASSOC);
        // liberazione delle risorse occupate dal risultato
        $result->close();
    }
}



$azienda = $row['azienda'];

if ($row['vedi_prezzi'] == 1) {
    $testo_prezzi = " , prezzi al " . date("d/m/y");
    $prezzi = true;
} else {
    $prezzi = false;
    $testo_prezzi = ", prezzi non visualizzati";
}



//$excel->getActiveSheet()->getRowDimension('A')->setVisible(FALSE);
// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$objPHPExcel->getActiveSheet()->getTabColor()->setRGB('973566');
$objPHPExcel->getDefaultStyle()->getFont()->setName('Calibri')->setSize(10);

$objPHPExcel->setActiveSheetIndex(0);
// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Ordine');

// Set document properties
$objPHPExcel->getProperties()->setCreator("La Rochelle")
        ->setLastModifiedBy("La Rochelle")
        ->setTitle("Modulo ordine precompilato")
        ->setSubject("Cliente " . $row['azienda'])
        ->setDescription("Modulo ordine precompilato " . $row['azienda'] . $testo_prezzi);
// ->setKeywords("office 2007 openxml php")
// ->setCategory("Test result file");
//Immagine
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Logo');
$objDrawing->setDescription('Logo');
$objDrawing->setPath('logo_documenti.jpg');
//$objDrawing->setResizeProportional(true);
$objDrawing->setWidthAndHeight(140, 100);

$objDrawing->setCoordinates('A1');
//$objDrawing->setOffsetX(110);
//$objDrawing->setRotation(25);
//$objDrawing->getShadow()->setVisible(true);
//$objDrawing->getShadow()->setDirection(45);
$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());

//Cliente
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('B1', 'Modulo ordine cliente');
$objPHPExcel->getActiveSheet()->getStyle('B1')->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$objPHPExcel->getActiveSheet()->mergeCells('B2:B3');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('B2', $row['azienda']);
$objPHPExcel->getActiveSheet()->getStyle('B2')->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
//Generato il
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('C1', 'Generato il');
$objPHPExcel->getActiveSheet()->getStyle('C1')->applyFromArray(array('font' => array('bold' => true, 'size' => 11)));
$objPHPExcel->getActiveSheet()->mergeCells('C2:C3');
$objPHPExcel->setActiveSheetIndex(0)->setCellValue('C2', date("d/m/y"));
$objPHPExcel->getActiveSheet()->getStyle('C2')->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$objPHPExcel->getActiveSheet()->getStyle('C2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);

//Intestazioni
$row = 5;



$objPHPExcel->getActiveSheet()
        ->setCellValue('A' . $row, 'Codice')
        ->setCellValue('B' . $row, 'Articolo')
        ->setCellValue('C' . $row, 'Prezzo');

$styleArray = array(
    'font' => array(
        'bold' => true, 'size' => 16
    )
);

$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('B' . $row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

$objPHPExcel->getActiveSheet()->getStyle('C' . $row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('C' . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);




//Recupero elenco dipendenti
$sql = "select dipendenti.*, misure.data_rilievo, utenti_gradi.nome_grado, utenti_gradi.colore, utenti_gradi.colore_nominativo FROM (dipendenti left JOIN utenti_gradi ON dipendenti.grado = utenti_gradi.id) left join (SELECT iddip,max(data_rilievo) as data_rilievo FROM dipendenti_misure group by iddip  ) as misure ON dipendenti.iddip = misure.iddip inner join utenti_dipendenti on dipendenti.iddip = utenti_dipendenti.iddip where utenti_dipendenti.dimesso=0 and utenti_dipendenti.iduser={$iduser}";

$col_dipendenti = 3;
$col = $col_dipendenti; //Colonna iniziale elenco dipendenti
if (!$dipendenti = $conn->query($sql)) {
    echo "Errore della query: " . $conn->error . ".";
    exit();
} else {
    // conteggio dei record
    $num_dipendenti = $dipendenti->num_rows;
    if ($num_dipendenti > 0) {
        // conteggio dei record restituiti dalla query
        while ($dipendente = $dipendenti->fetch_array(MYSQLI_ASSOC)) {

            // Add some data
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $dipendente['cognome'] . " " . $dipendente['nome']);
            
            $cellaString=PHPExcel_Cell::stringFromColumnIndex($col) . $row;

            //Ruota testo 90°
            $objPHPExcel->getActiveSheet()->getStyle($cellaString)->getAlignment()->setTextRotation(90);
            //Allineamento orizzontale centrato
            $objPHPExcel->getActiveSheet()->getStyle($cellaString)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



            $keyCell = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row);



            //$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($col, $row)->getAlignment()->setTextRotation(90);
            //	$objPHPExcel->getActiveSheet()->getColumnDimension($col)->setAutoSize(true);
		//$objPHPExcel->getActiveSheet()->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($col))->setAutoSize(true);



            $col++;
        }
        // liberazione delle risorse occupate dal risultato
        $dipendenti->close();
    }
}
if ($num_dipendenti > 0) {
    //Aggiungo quantita totale
    // Add some data
    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Quantità totale');

    //Ruota testo 90°
    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->getAlignment()->setTextRotation(90);
    //Allineamento orizzontale centrato
    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->applyFromArray($styleArray);
    $col_tot_dipendenti = $num_dipendenti + 10;
} else {
    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Quantità');

    //Ruota testo 90°
    //$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col).$row)->getAlignment()->setTextRotation(90);
    //Allineamento orizzontale centrato
    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->applyFromArray($styleArray);
    $objPHPExcel->getActiveSheet()->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($col))->setAutoSize(true);
}

$col++;
$col_totale = $col;
//Aggiungo importo totale
// Add some data
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, 'Totale');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
//Ruota testo 90°
//$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col).$row)->getAlignment()->setTextRotation(90);
//Allineamento orizzontale centrato
//$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col).$row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col) . $row)->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($col))->setAutoSize(true);






$styleBorderArray = array(
    'borders' => array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000'),
        ),
        'inside' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000'),
        )
    ),
    'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => 'C8C8C8'))
);
$styleBorderArray2 = array(
    'borders' => array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000'),
        ),
        'inside' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000'),
        )
    )
);
$styleBorderArray3 = array(
    'borders' => array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => '000000'),
        )
    )
);




$editabile = array(
    'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array('rgb' => '99e6ff')));





//Sfondo prima riga tabella
$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col) . $row)
        ->applyFromArray($styleBorderArray);



//Recupero articoli nel carrello
$sql = "select preferiti.*, prodotti.codice, prodotti.articolo, prodotti.prezzo, prodotti.costo, prodotti.Promozione, prodotti.Sconto, prodotti.Prodata,  prodotti.idfor, prodotti.fileimg, prodotti.aggiungi_a_ordine FROM prodotti left JOIN preferiti  ON preferiti.idpro = prodotti.IDpro where iduser={$iduser}";

$col = 1; //Colonna iniziale elenco dipendenti
$row++;
if (!$articoli = $conn->query($sql)) {
    echo "Errore della query: " . $conn->error . ".";
    exit();
} else {
    $num_articoli = $articoli->num_rows;

    if ($num_articoli > 0) {
        // conteggio dei record restituiti dalla query
        while ($articolo = $articoli->fetch_array(MYSQLI_ASSOC)) {

            // Add some data
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, $articolo['codice']);

            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1, $row, utf8_encode($articolo['articolo']));

            if ($prezzi) {
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, $articolo['prezzo']);
            }
            $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(2) . $row)->getNumberFormat()->setFormatCode('0.00');

            if ($num_dipendenti > 0) {
                //Aggiungo formula per totale quantità dipendenti
                $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3 + $num_dipendenti, $row, '=SUM(D' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex(2 + $num_dipendenti) . $row . ')');
                $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(3 + $num_dipendenti) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
                //Tolgo protezione a quantita dipendenti
                $objPHPExcel->getActiveSheet()->getStyle('D' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex(2 + $num_dipendenti) . $row . '')->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED);
                $objPHPExcel->getActiveSheet()->getStyle('D' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex(2 + $num_dipendenti) . $row . '')->applyFromArray($editabile);
                //Ciclo per colonne parziali dipendente
                for ($col = 0; $col < $num_dipendenti; $col++) {
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col + $col_tot_dipendenti, $row, '=' . PHPExcel_Cell::stringFromColumnIndex(2) . $row . '*' . PHPExcel_Cell::stringFromColumnIndex($col + $col_dipendenti) . ($row));

                    $objPHPExcel->getActiveSheet()->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($col + $col_tot_dipendenti))->setVisible(FALSE);
                }
            } else {
                //Tolgo protezione a quantita
                $objPHPExcel->getActiveSheet()->getStyle('D' . $row)->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED);
                $objPHPExcel->getActiveSheet()->getStyle('D' . $row)->applyFromArray($editabile);
            }
            //Aggiungo formula per totale prezzo
            $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col_totale, $row, '=' . PHPExcel_Cell::stringFromColumnIndex(3 + $num_dipendenti) . $row . '*C' . $row . ')');
            $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
            $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->getNumberFormat()->setFormatCode('0.00');
            $row++;
        }
        // liberazione delle risorse occupate dal risultato
        $articoli->close();
    }
}


//Riga trasporto
//$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, $articolo['codice']);
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1, $row, 'Trasporto');

if ($prezzi) {
    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, 12.00);
}
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(2) . $row)->getNumberFormat()->setFormatCode('0.00');

//Aggiungo formula per totale prezzo
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col_totale, $row, 12.00);
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->getNumberFormat()->setFormatCode('0.00');
$row++;



//Bordi righe articoli
$objPHPExcel->getActiveSheet()
        ->getStyle('A6:' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row - 1))
        ->applyFromArray($styleBorderArray2);

//Riga totale dipendenti
if ($num_dipendenti && $prezzi) {
    $objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':C' . $row); //Merge
    $objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray($styleBorderArray); //Bordi
    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Totale per dipendente iva esclusa');
    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));

//for($i=$col_dipendenti;$i<=($col_dipendenti+$num_dipendenti);$i++){
//	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($i, $row, '=SUM('.PHPExcel_Cell::stringFromColumnIndex($i+ $col_tot_dipendenti).'6:'.PHPExcel_Cell::stringFromColumnIndex($i + $col_tot_dipendenti).($row-2).')');
//}
    for ($i = 0; $i < $num_dipendenti; $i++) {
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($i + $col_dipendenti, $row, '=SUM(' . PHPExcel_Cell::stringFromColumnIndex($i + $col_tot_dipendenti) . '6:' . PHPExcel_Cell::stringFromColumnIndex($i + $col_tot_dipendenti) . ($row - 2) . ')');
        $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($i + $col_dipendenti) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
        $objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($i + $col_dipendenti) . $row)->getNumberFormat()->setFormatCode('0.00');
    }
    $row++;
}









//Riga totale
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale - 1) . $row); //Merge
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray($styleBorderArray); //Bordi
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Totale iva esclusa');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col_totale, $row, '=SUM(' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . '6:' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row - 1) . ')');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->getNumberFormat()->setFormatCode('0.00');
//Riga iva
$row++;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale - 1) . $row); //Merge
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray($styleBorderArray); //Bordi
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Iva 22%');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col_totale, $row, '=round(' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row - 1) . '*0.22,2)');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->getNumberFormat()->setFormatCode('0.00');
//Riga totale con iva
$row++;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale - 1) . $row); //Merge
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray($styleBorderArray); //Bordi
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Totale con iva');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col_totale, $row, '=' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row - 2) . '+' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row - 1));
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 16)));
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col_totale) . $row)->getNumberFormat()->setFormatCode('0.00');



$row += 2;
//nota compilazione
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':B' . ($row));
$objPHPExcel->getActiveSheet()->setCellValue('A' . $row, 'Indicare con X se TRATTATIVA DIRETTA MePA o ORDINE DIRETTO FUORI MePA');
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13, 'color' => array('rgb' => 'FF0000'))));
$sheet = $objPHPExcel->getActiveSheet();
$sheet->getStyle('A' . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);


$row++;
//TRATTATIVA DIRETTA MePA con PREZZO A CORPO

$objPHPExcel->getActiveSheet()->setCellValue('B' . $row, 'TRATTATIVA DIRETTA MePA con PREZZO A CORPO');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(1) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));

$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row)
        ->applyFromArray($styleBorderArray2);
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED); //Tolgo protezione
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray(array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER), 'font' => array('bold' => true, 'size' => 15)));
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray($editabile);

//ORDINE DIRETTO FUORI MePA
$row++;
$objPHPExcel->getActiveSheet()->setCellValue('B' . $row, 'ORDINE DIRETTO FUORI MePA');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(1) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row)
        ->applyFromArray($styleBorderArray2);
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED); //Tolgo protezione
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray(array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER), 'font' => array('bold' => true, 'size' => 15)));
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray($editabile);

//Avviso TRATTATIVA DIRETTA MePA con PREZZO A CORPO
$row ++;
$objPHPExcel->getActiveSheet()->setCellValue('B' . $row, '=IF(A'.($row-2).'<>"","Concludere la procedura di stipula entro la data da voi fissata, altrimenti dovrete eseguire nuovamente il caricamento su MePA.","")');
$objPHPExcel->getActiveSheet()->getStyle('B' . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 12, 'color' => array('rgb' => 'FF0000'))));



//Note
$row ++;
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Note');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$row++;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row + 3));
$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row + 3))
        ->applyFromArray($styleBorderArray2);
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . 'A' . ($row + 3))->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED); //Tolgo protezione
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . 'A' . ($row + 3))->applyFromArray($editabile);

$row += 3;

//Altri dati
$row += 2;
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'CIG');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$row++;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row));
$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row))
        ->applyFromArray($styleBorderArray2);
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . 'A' . ($row))->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED); //Tolgo protezione
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . 'A' . ($row))->applyFromArray($editabile);


$row += 2;
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Determinazione');
$objPHPExcel->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex(0) . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));
$row++;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row));
$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row))
        ->applyFromArray($styleBorderArray2);
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . 'A' . ($row))->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED); //Tolgo protezione
$objPHPExcel->getActiveSheet()->getStyle('A' . $row . ':' . 'A' . ($row))->applyFromArray($editabile);

//Annotazioni in rosso	
$row += 2;
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'Attenzione:');
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13, 'color' => array('rgb' => 'FF0000'))));

$objPHPExcel->getActiveSheet()->getStyle('B' . $row)->getAlignment()->setVertical('top');
//$objPHPExcel->getActiveSheet()->mergeCells('B'.$row.':'.PHPExcel_Cell::stringFromColumnIndex($col_totale).($row));
$objPHPExcel->getActiveSheet()->getStyle('B' . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 11, 'color' => array('rgb' => 'FF0000'))));
$objPHPExcel->getActiveSheet()->setCellValue('B' . $row, "1 ) Nel trasmettere la richiesta di fornitura allegare copia della DETERMINAZIONE.\n2 ) Comunicare se rispetto all’ultimo ordine trasmesso sono necessarie delle modifiche nelle taglie.\n3 ) Nel caso di ordine perfezionato con TRATTATIVA DIRETTA MePA dovete concludere la procedura di stipula\n entro la data da voi fissata, altrimenti dovrete eseguire nuovamente il caricamento su MePA.");
$objPHPExcel->getActiveSheet()->getStyle('B' . $row)->getAlignment()->setWrapText(true);

$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(-1);




//Dati La Rochelle
$row += 2;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row));
$objPHPExcel->getActiveSheet()
        ->getStyle('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row + 1))
        ->applyFromArray($styleBorderArray3);
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->setCellValue('A' . $row, 'Laboratorio Artigiano di Pistono & C. S.N.C., Via Monsignor A. Sangiorgio 59 - 10090 San Giorgio Canavese (TO)');
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->applyFromArray(array('font' => array('bold' => true, 'size' => 13)));

$row++;
$objPHPExcel->getActiveSheet()->mergeCells('A' . $row . ':' . PHPExcel_Cell::stringFromColumnIndex($col_totale) . ($row));
$objPHPExcel->getActiveSheet()->getStyle('A' . $row)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, 'P.IVA 06315200011, Licenza EX. ART. 28 T.U.L.P.S. - Prot. 0077320/2014 AREA  I TER');





//Blocco roga colonna
if ($num_dipendenti > 0) {

$objPHPExcel->getActiveSheet()->freezePane('C6');
}

//Protezione foglio
$objPHPExcel->getActiveSheet()->getProtection()->setSheet(true);
$objPHPExcel->getActiveSheet()->getProtection()->setSort(true);
//$objPHPExcel->getActiveSheet()->getProtection()->setInsertRows(true);
$objPHPExcel->getActiveSheet()->getProtection()->setFormatCells(true);

$objPHPExcel->getActiveSheet()->getProtection()->setPassword('password');








// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
header('Content-Disposition: attachment;filename="Modulo_Ordine.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');



$txt_log = "[utente={$iduser}]{$azienda}[/utente] ha scaricato ordine excel\r\nDipendenti:{$num_dipendenti}<br>Articoli:{$num_articoli}<br>" . $testo_prezzi . "<br>Scaricato da:{$data['iduser']}";
$txt_log = $conn->real_escape_string($txt_log);
$sql = "insert into log (data, evento, causale) values (now(), '{$txt_log}', 1)";
$conn->query($sql);


exit;



