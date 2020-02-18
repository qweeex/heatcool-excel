<?php
    header('Content-Type: text/html; charset=utf8');
    // Main Class Excel
    require_once 'lib/PHPExcel.php';
    // Excel Class writer
    require_once 'lib/PHPExcel/Writer/Excel5.php';


    $xls = new PHPExcel();

    $xls->setActiveSheetIndex(0);

    $sheet = $xls->getActiveSheet();
    $sheet->setTitle('list 1');


    // Inside text

    $sheet->setCellValue("A1", 'Zálohová faktúra č.200400001');


    // Объединение ячеек
    $sheet->mergeCells('A1:N1');

    // Text right

    $sheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);



    // Выводим содержимое файла
    $objWriter = new PHPExcel_Writer_Excel5($xls);
    $objWriter->save('php://output');


    echo 'good';