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
    $sheet->setCellValue("D12", "VÚB Banka, a.s.");

