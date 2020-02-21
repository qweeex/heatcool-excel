<?php

// Main Class Excel
require_once 'lib/PHPExcel.php';
// Excel Class writer
require_once 'lib/PHPExcel/Writer/Excel5.php';
$xls = new PHPExcel();
$xls->setActiveSheetIndex(0);

$sheet = $xls->getActiveSheet();
$sheet->setTitle('list 1');


/// Config

$sheet->getColumnDimension('G')->setWidth(25);
$sheet->getColumnDimension('I')->setWidth(25);


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

$sheet->setCellValue("I18", "Konečný príjemca:");
$sheet->mergeCells('I18:M18');
$sheet->getStyle("I18")->getFont()->setBold(true);

$sheet->mergeCells('I19:M23');


// Заголовки таблиц

$sheet->setCellValue("B25", "Označenie dodávky");
$sheet->mergeCells('B25:D25');
$sheet->getStyle("B25")->getFont()->setBold(true);

$sheet->setCellValue("E25", "Množstvo");
$sheet->mergeCells('E25:F25');
$sheet->getStyle("E25")->getFont()->setBold(true);
$sheet->getStyle("E25")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("G25", "J.cena");
$sheet->getStyle("G25")->getFont()->setBold(true);
$sheet->getStyle("G25")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("H25", "Cena");
$sheet->mergeCells('H25:I25');
$sheet->getStyle("H25")->getFont()->setBold(true);
$sheet->getStyle("H25")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("J25", "%DPH");
$sheet->getStyle("J25")->getFont()->setBold(true);
$sheet->getStyle("J25")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("K25", "DPH");
$sheet->getStyle("K25")->getFont()->setBold(true);
$sheet->getStyle("K25")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue("L25", "EUR Celkom");
$sheet->mergeCells('L25:M25');
$sheet->getStyle("L25")->getFont()->setBold(true);
$sheet->getStyle("L25")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


// Заголовок перед списком товаров

$sheet->setCellValue("B28", "Fakturujeme Vám zálohu za objednaný tovar:  ");
$sheet->mergeCells('B28:M28');
$sheet->getStyle("B28")->getFont()->setBold(true);


// Список товаров

$count = 31;

$arr = array(
    array(
        'name' => 'NIBE SPLIT - SET 5',
        'count' => 1,
        'price-j' => '6500',
        'price' => '6500',
        'dph_plus' => '0',
        'dph' => '0',
        'eur' => '6500'
    ),
    array(
        'name' => 'NIBE SPLIT - SET 5',
        'count' => 1,
        'price-j' => '6500',
        'price' => '6500',
        'dph_plus' => '0',
        'dph' => '0',
        'eur' => '6500'
    ),
    array(
        'name' => 'NIBE SPLIT - SET 5',
        'count' => 1,
        'price-j' => '6500',
        'price' => '6500',
        'dph_plus' => '0',
        'dph' => '0',
        'eur' => '6500'
    )
);

foreach ($arr as $item) {

    $sheet->setCellValue("B" .$count , $item['name']);
    $sheet->mergeCells('B'.$count.':D'.$count);

    $sheet->setCellValue("E" .$count, $item['count']);
    $sheet->mergeCells('E'.$count.':F'.$count);
    $sheet->getStyle("E31")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $sheet->setCellValue("G" . $count, $item['price-j']);
    $sheet->getStyle("G" . $count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $sheet->setCellValue("H" . $count, $item['price']);
    $sheet->mergeCells('H'.$count.':I' . $count);
    $sheet->getStyle("H31")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $sheet->setCellValue("J" . $count, $item['dph_plus']);
    $sheet->getStyle("J" . $count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $sheet->setCellValue("K" . $count, $item['dph']);
    $sheet->getStyle("K" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $sheet->setCellValue("L" .$count, $item['eur']);
    $sheet->mergeCells('L'.$count.':M' .$count);
    $sheet->getStyle("L" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

    $count++;
    $count++;
}


// Итоговая цена

$count++;

$sheet->setCellValue("B" .$count , "Súčet položiek");
$sheet->mergeCells('B'.$count.':M'.$count);

$count++;

$sheet->setCellValue("B" .$count , "SPOLU NA ÚHRADU");
$sheet->mergeCells('B'.$count.':C'.$count);
$sheet->getStyle("B" .$count)->getFont()->setBold(true);


// Сама цена
$sheet->setCellValue("J" .$count , "6481,2");
$sheet->mergeCells('J'.$count.':M'.$count);
$sheet->getStyle("J" .$count)->getFont()->setBold(true);
$sheet->getStyle("J" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$count++;
$count++;

// Информация
$sheet->setCellValue("B" .$count , "Vystavil:");
$sheet->mergeCells('B'.$count.':C'.$count);

$sheet->setCellValue('D'.$count, 'systém obchodtepelnecerpadla');

$count++;
$count++;
$count++;

/// Текст

$sheet->setCellValue("B" .$count , "Dovoľujeme si Vás upozorniť, že v prípade nedodržania termínu splatnosti uvedeného na faktúre, Vám môžeme účtovať úrok z omeškania v dohodnutej, resp. zákonnej výške a zmluvnú pokutu (ak bola dohodnutá).");
$count2 = $count + 1;
$sheet->mergeCells('B'.$count.':M'.$count2);
$sheet->getStyle('B'.$count)->getAlignment()->setWrapText(true);

$count++;
$count++;
$count++;


/// Налоги

$sheet->setCellValue('B'.$count, "Rekapitulácia v EUR:");
$sheet->mergeCells("B".$count.":D".$count);
$sheet->getStyle("B".$count)->getFont()->setBold(true);

$sheet->setCellValue('E'.$count, "Základ v EUR");
$sheet->mergeCells("E".$count.":G".$count);
$sheet->getStyle("E".$count)->getFont()->setBold(true);
$sheet->getStyle("E" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('H'.$count, "Sadzba");
$sheet->mergeCells("H".$count.":I".$count);
$sheet->getStyle("H".$count)->getFont()->setBold(true);
$sheet->getStyle("H" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('J'.$count, "DPH v EUR");
$sheet->mergeCells("J".$count.":K".$count);
$sheet->getStyle("J".$count)->getFont()->setBold(true);
$sheet->getStyle("J" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('L'.$count, "Spolu s DPH v EUR");
$sheet->mergeCells("L".$count.":M".$count);
$sheet->getStyle("L".$count)->getFont()->setBold(true);
$sheet->getStyle("L" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$count++;
$count++;
$count++;

// Заполняем налоги

$sheet->setCellValue('B'.$count, "No info");
$sheet->mergeCells("B".$count.":D".$count);

$sheet->setCellValue('E'.$count, "10");
$sheet->mergeCells("E".$count.":G".$count);
$sheet->getStyle("E" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('H'.$count, "10%");
$sheet->mergeCells("H".$count.":I".$count);
$sheet->getStyle("H" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('J'.$count, "");
$sheet->mergeCells("J".$count.":K".$count);
$sheet->getStyle("J" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$sheet->setCellValue('L'.$count, "10");
$sheet->mergeCells("L".$count.":M".$count);
$sheet->getStyle("L" .$count)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$count++;
$count++;
$count++;


/// Подвал

$sheet->setCellValue('B'.$count, "Prevzal, podpis:");
$sheet->mergeCells("B".$count.":C".$count);

$sheet->setCellValue('H'.$count, "Vystavil, podpis:");
$sheet->mergeCells("H".$count.":I".$count);

$count++;
$count++;
$count++;

$sheet->setCellValue("B".$count, "Ekonomický a informačný systém POHODA ");
$sheet->mergeCells("B".$count.":F".$count);


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