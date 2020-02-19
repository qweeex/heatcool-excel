<?php

// Main Class Excel
require_once 'lib/PHPExcel.php';
// Excel Class writer
require_once 'lib/PHPExcel/Writer/Excel5.php';
$xls = new PHPExcel();
$xls->setActiveSheetIndex(0);

$sheet = $xls->getActiveSheet();
$sheet->setTitle('list 1');


// Шапка таблицы
$sheet->setCellValue("A1", "Zálohová faktúra č.200400001");
$sheet->mergeCells('A1:N1');
$sheet->getStyle("A1")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


// Информация о компании
$sheet->setCellValue("B3", "Dodávateľ:");
$sheet->setCellValue("B4", "Heatcool s.r.o.");
$sheet->setCellValue("B5", "Šancová 50");
$sheet->setCellValue("B6", "811 05  Bratislava - mestská časť Staré Mesto");
$sheet->getStyle("B4")->getFont()->setBold(true);
$sheet->getStyle("B5")->getFont()->setBold(true);
$sheet->getStyle("B6")->getFont()->setBold(true);

// Реквизиты
$sheet->setCellValue("B8", "IČO: 50 198 637");
$sheet->setCellValue("B9", "DIČ: 2120219882");
$sheet->setCellValue("B10", "IČ DPH: SK2120219882");


// Банковские реквизиты
$sheet->setCellValue("B12", "Banka:");
$sheet->setCellValue("D12", "VÚB Banka, a.s.");

$sheet->setCellValue("B13", "SWIFT:");

$sheet->setCellValue("B14", "IBAN:");
$sheet->setCellValue("D14", "SK82 0200 0000 0039 3866 7854");

$sheet->setCellValue("B15", "Číslo účtu:");
$sheet->setCellValue("D15", "#");

$sheet->setCellValue("B16", "Kód banky:");
$sheet->setCellValue("D16", "#");


// Информация о заказе
$sheet->setCellValue("I3", "Variabilný symbol:");
$sheet->mergeCells('I3:J3');
$sheet->setCellValue("K3", "200400001");
$sheet->mergeCells('K3:M3');
$sheet->getStyle("K3")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


$sheet->setCellValue("I4", "Konštantný symbol:");
$sheet->mergeCells('I4:J4');
$sheet->setCellValue("K4", "308");
$sheet->mergeCells('K4:M4');
$sheet->getStyle("K4")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


$sheet->setCellValue("I5", "Objednávka č.");
$sheet->mergeCells('I5:J5');
$sheet->setCellValue("K5", "zo dňa ");
$sheet->mergeCells('K5:M5');
$sheet->getStyle("K5")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

// Информация о клиенте

$sheet->setCellValue("I7", "Odberateľ:");
$sheet->mergeCells('I7:M7');
$sheet->getStyle("I7")->getFont()->setBold(true);

$sheet->setCellValue("I8" ,"IČO:");
$sheet->setCellValue("J8", "54654654");
$sheet->mergeCells('J8:M8');
$sheet->getStyle("J8")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("I9" ,"DIČ:");
$sheet->setCellValue("J9", "54654654");
$sheet->mergeCells('J9:M9');
$sheet->getStyle("J9")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("I10" ,"IČ DPH:");
$sheet->setCellValue("J10", "546540656");
$sheet->mergeCells('J10:M10');
$sheet->getStyle("J10")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("I11" ,"Meno");
$sheet->setCellValue("J11", "Ivan Ivanov Ivanic");
$sheet->mergeCells('J11:M11');
$sheet->getStyle("J11")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("I12" ,"Ulica a číslo domu");
$sheet->setCellValue("J12", "Saratov, Komsomol, 50m, 60");
$sheet->mergeCells('J12:M12');
$sheet->getStyle("J12")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("I13" ,"PSČ a obec/mesto");
$sheet->setCellValue("J13", "100");
$sheet->mergeCells('J13:M13');
$sheet->getStyle("J13")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


// Даты

$sheet->setCellValue("B18", "Dátum vyhotovenia:");
$sheet->mergeCells('B18:D18');

$sheet->setCellValue("G18", "19.02.2020");
$sheet->getStyle("G18")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("B19", "Dátum splatnosti:");
$sheet->mergeCells('B19:D19');

$sheet->setCellValue("G19", "19.02.2020");
$sheet->getStyle("G19")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("B20", "Dátum dodania tovaru/služby; prijatie platby:");
$sheet->mergeCells('B20:F20');

$sheet->setCellValue("G20", "19.02.2020");
$sheet->getStyle("G20")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("B21", "Forma úhrady:");
$sheet->mergeCells('B21:D21');

$sheet->setCellValue("E21", "19.02.2020");
$sheet->mergeCells('E21:G21');
$sheet->getStyle("E21")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


// Комментарии

$sheet->setCellValue("I18", "Odberateľ:");
$sheet->mergeCells('I18:M18');
$sheet->getStyle("I18")->getFont()->setBold(true);

$sheet->mergeCells('I19:M23');


// Заголовки таблиц


// Выводим HTTP-заголовки
header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
header ( "Cache-Control: no-cache, must-revalidate" );
header ( "Pragma: no-cache" );
header ( "Content-type: application/vnd.ms-excel" );
header ( "Content-Disposition: attachment; filename=matrix.xls" );

// Выводим содержимое файла
$objWriter = new PHPExcel_Writer_Excel5($xls);
$objWriter->save('php://output');