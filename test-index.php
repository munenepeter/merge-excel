<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

ini_set('memory_limit', '-1'); //# might over use the memory, so remove the limit
date_default_timezone_set('Africa/Nairobi');

require 'vendor/autoload.php';

const UPLOAD_DIR = __DIR__ . '/uploads';
const OUTPUT_DIR = __DIR__ . '/output';
const FILE_TYPE = 'Xls';

$files = array_diff(scandir(UPLOAD_DIR), ['..', '.']);
$workbooks = array_map(fn ($file) => UPLOAD_DIR . DIRECTORY_SEPARATOR . $file, $files);

// Create a new spreadsheet to hold the merged workbooks
$merged_spreadsheet = new Spreadsheet();

// $merged_spreadsheet->getProperties()
//     ->setCreator("Peter")
//     ->setLastModifiedBy("Peter")
//     ->setTitle("Office 2007 XLSX Test Document")
//     ->setSubject("Office 2007 XLSX Test Document")
//     ->setDescription(
//         "Test document for Office 2007 XLSX, generated using PHP classes."
//     )
//     ->setKeywords("office 2007 openxml php")
//     ->setCategory("Test result file");

$sheet_index = 0;
$total_sheets = 0;
$stats = [];

// Create a new sheet for stats
$stats_sheet = $merged_spreadsheet->getActiveSheet();
$stats_sheet->setTitle('Merge Stats');

foreach ($workbooks as $workbook) {

    $spreadsheet = IOFactory::load($workbook);
    $sheet_count = $spreadsheet->getSheetCount();
    $total_sheets += $sheet_count;
    $stats[] = ['filename' => basename($workbook), 'sheet_count' => $sheet_count];

    // Loop through each sheet in the current workbook
    foreach ($spreadsheet->getAllSheets() as $sheet) {
        // Create a new sheet in the merged workbook
        $new_sheet = new Worksheet($merged_spreadsheet, basename($workbook, '.' . pathinfo($workbook, PATHINFO_EXTENSION)) . ' S' . $sheet_index);
        $merged_spreadsheet->addSheet($new_sheet);

        // Copy content from the source sheet to the new sheet
        copy_sheet_content($sheet, $new_sheet);

        $sheet_index++;
    }
}


// Fill stats sheet with information
$stats_sheet->setCellValue('A1', 'Merged ' . count($workbooks) . ' workbooks with a total of ' . $total_sheets . ' sheets');

$row = 3;
foreach ($stats as $stat) {
    $stats_sheet->setCellValue('A' . $row, $stat['filename']);
    $stats_sheet->setCellValue('B' . $row, $stat['sheet_count'] . ' sheets');
    $row++;
}

// Save the merged workbook
$writer = IOFactory::createWriter($merged_spreadsheet, 'Xlsx');
$output_file = OUTPUT_DIR . DIRECTORY_SEPARATOR . date('YmdHis') . '-output.xlsx';
$writer->save($output_file);

echo "Merged workbook created successfully!";

function copy_sheet_content(Worksheet $source_sheet, Worksheet $target_sheet) {
    foreach ($source_sheet->getRowIterator() as $row) {
        foreach ($row->getCellIterator() as $cell) {
            $coordinate = $cell->getCoordinate();
            $target_cell = $target_sheet->getCell($coordinate);
            $target_cell->setValue($cell->getValue());

            // Copy basic cell styles (e.g., font, alignment, fill)
            $source_style = $cell->getStyle();
            $target_style = $target_cell->getStyle();

            // Copy font
            $target_style->getFont()->setName($source_style->getFont()->getName());
            $target_style->getFont()->setSize($source_style->getFont()->getSize());
            $target_style->getFont()->setBold($source_style->getFont()->getBold());
            $target_style->getFont()->setItalic($source_style->getFont()->getItalic());
            $target_style->getFont()->getColor()->setARGB($source_style->getFont()->getColor()->getARGB());

            // Copy alignment
            $target_style->getAlignment()->setHorizontal($source_style->getAlignment()->getHorizontal());
            $target_style->getAlignment()->setVertical($source_style->getAlignment()->getVertical());

            // Copy fill
            $target_style->getFill()->setFillType($source_style->getFill()->getFillType());
            $target_style->getFill()->getStartColor()->setARGB($source_style->getFill()->getStartColor()->getARGB());

            //borders
            $target_style->getBorders()->getTop()->setBorderStyle($source_style->getBorders()->getTop()->getBorderStyle());
            $target_style->getBorders()->getBottom()->setBorderStyle($source_style->getBorders()->getBottom()->getBorderStyle());
            $target_style->getBorders()->getLeft()->setBorderStyle($source_style->getBorders()->getLeft()->getBorderStyle());
            $target_style->getBorders()->getRight()->setBorderStyle($source_style->getBorders()->getRight()->getBorderStyle());
        }
    }

    // Copy merged cells
    foreach ($source_sheet->getMergeCells() as $merged_cell_range) {
        $target_sheet->mergeCells($merged_cell_range);
    }
    //auto resize cols
    foreach (range('A', $target_sheet->getHighestColumn()) as $col) {
        $target_sheet->getColumnDimension($col)->setAutoSize(true);
    }

    //auto resize rows see - https://stackoverflow.com/questions/46578567/phpspreadsheet-auto-row-height-not-works-with-libreoffice-latest-version
    foreach (range(1, $target_sheet->getHighestRow()) as $row) {
        $target_sheet->getRowDimension($row)->setRowHeight(-1);
    }
}
