<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('America/Caracas');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once ('/Classes/PHPExcel.php');

//echo $resul;
$objPHPExcel = new PHPExcel();
$objPHPExcel2 = new PHPExcel();


$objPHPExcel->
  getProperties()
    ->setCreator("ORINOCONET")
    ->setTitle("Matriz Ambiental")
    ->setSubject("Documento")
    ->setDescription("Documento generado con PHPExcel");

 $borders = array(
          'borders' => array(
            'allborders' => array(
              'style' => PHPExcel_Style_Border::BORDER_THICK,
              'color' => array('argb' => 'FFA5A5A5'),
            )
          ),
        );
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setAutoSize(true);

$objPHPexcel2 = PHPExcel_IOFactory::load('Formato_reporte_SAO.xlsx');
$objClonedWorksheet = clone $objPHPexcel2->getSheetByName('Hoja1');
$objPHPExcel->addExternalSheet($objClonedWorksheet);
// Add some data
// $sheetIndex = $objPHPExcel->getIndex($objPHPExcel-> getSheetByName('Worksheet'));
// $objPHPExcel->removeSheetByIndex($sheetIndex);
// $sheetIndex = $objPHPExcel->getIndex($objPHPExcel-> getSheetByName('Hoja1'));
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('F13', $var[0]['proceso']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('L13', $var[0]['subproceso']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('S13', $var[0]['fecha_emision']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('Y13', $var[0]['fecha_actualizacion']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('AC13', $var[0]['contador']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('C11', 'REF:  '.$var[0]['ref']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('E8', $var[0]['nombre_unidad']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('AC9', $var[0]['vigencia']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('AA7', $var[0]['codigo_form']);
// $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue('AC11', $var[0]['fecha_emisionF']);
//
// $casilla='16';
// for ($cont=1,$i=0; $i < $tam; $i++) {
//   $aux=(int)$casilla;
//   $aux++;
//   $casilla=(string)$aux;
//   $objPHPExcel->getActiveSheet()->getStyle('C'.$casilla.':AC'.$casilla)->applyFromArray($borders);
//   $cell='C'.$casilla;
//   if ($i!=0) {
//     if ( $var[$i]['actv_prod_serv']!= $var[$i-1]['actv_prod_serv']) {
//       $cont++;
//     }
//   }
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $cont);
//   $cell='D'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['actv_prod_serv']);
//   $cell='E'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['descripcion']);
//   $cell='F'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['medidas']);
//   $cell='G'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['aspecto']);
//   if ($var[$i]['condicion']==="NORMAL") {
//     $cell='H'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, 'X');
//   }
//   if ($var[$i]['condicion']==="EMERGENCIA") {
//     $cell='J'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, 'X');
//   }
//   if ($var[$i]['condicion']==="ANORMAL") {
//     $cell='I'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, 'X');
//   }
//   $cell='K'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['impacto']);
//   $cell='L'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['receptor']);
//   $cell='N'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['frecuencia']);
//   $cell='O'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['severidad']);
//   $cell='P'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['alcance']);
//   $cell='Q'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['crit_impacto_ambiental']);
//   $cell='R'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['existencia']);
//   $cell='S'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['cumplimiento']);
//   $cell='T'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['cltotal']);
//   $cell='U'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['existencia']);
//   $cell='V'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['gestion']);
//   $cell='W'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['partes_interesadas']);
//   $cell='X'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['ei_res_evaluacion']);
//   $cell='Y'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['significacion']);
//   $cell='Z'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['opciones_tec']);
//   $cell='AA'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['opciones_econ']);
//   $cell='AC'.$casilla;
//   $objPHPExcel->setActiveSheetIndex($sheetIndex)
//               ->setCellValue($cell, $var[$i]['observaciones']);
// }



$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));

// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// $objWriter->save('php://output');
exit;

//$objPHPExcel->getActiveSheet()->setCellValue('A8',"Hello\nWorld");
//$objPHPExcel->getActiveSheet()->getRowDimension(8)->setRowHeight(-1);
//$objPHPExcel->getActiveSheet()->getStyle('A8')->getAlignment()->setWrapText(true);

//$objPHPExcel->getActiveSheet()->setTitle('Simple');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet


// Save Excel 2007 file


//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save('php://output');
