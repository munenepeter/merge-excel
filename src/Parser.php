<?php

namespace MergeExcel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use ZipStream\Exception\FileNotFoundException;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Parser {

    private $new_spreadsheet = null;
    private $total_sheets = 0;
    private $no_of_workbooks = 0;

    public function __construct() {
        $this->create_workbook();
    }

    public function __destruct() {
        $this->new_spreadsheet = null;
    }

    public function convert(string $workbook, string $from = '', string $to = 'xlsx', callable $progressCallback = null): bool {

        if (!$this->new_spreadsheet instanceof Spreadsheet)
            throw new \LogicException("We have not created a new spreadsheet!");

        $sheet_index = 0;
        $stats = [];

        $spreadsheet = IOFactory::load($workbook);
        $sheet_count = $spreadsheet->getSheetCount();
        $this->total_sheets += $sheet_count;

        $stats[] = ['filename' => basename($workbook), 'sheet_count' => $sheet_count];

        foreach ($spreadsheet->getAllSheets() as $sheet) {

            $new_sheet = new Worksheet($this->new_spreadsheet, basename($workbook, '.' . pathinfo($workbook, PATHINFO_EXTENSION)) . ' S' . $sheet_index);

            $this->new_spreadsheet->addSheet($new_sheet);

            $this->copy_sheet_content($sheet, $new_sheet);
        }
        if ($progressCallback) {
            $progressCallback($sheet_index);
        }

        $sheet_index++;

        $this->write_stats(operation: "converted", statistics: $stats);

        return true;
    }

    public function merge(string $path, callable $progressCallback = null): bool {

        if (!$this->new_spreadsheet instanceof Spreadsheet)
            throw new \LogicException("We have not created a new spreadsheet!");

        $sheet_index = 0;
        $stats = [];

        foreach ($this->get_workbooks($path) as $workbook) {

            $spreadsheet = IOFactory::load($workbook);
            $sheet_count = $spreadsheet->getSheetCount();
            $this->total_sheets += $sheet_count;

            $stats[] = ['filename' => basename($workbook), 'sheet_count' => $sheet_count];

            foreach ($spreadsheet->getAllSheets() as $sheet) {

                $new_sheet = new Worksheet($this->new_spreadsheet, basename($workbook, '.' . pathinfo($workbook, PATHINFO_EXTENSION)) . ' S' . $sheet_index);

                $this->new_spreadsheet->addSheet($new_sheet);

                $this->copy_sheet_content($sheet, $new_sheet);
            }

            $sheet_index++;
            
            if ($progressCallback) {
                $progressCallback($this->no_of_workbooks);
            }
        }

        $this->write_stats(operation: "merged", statistics: $stats);

        return true;
    }

    private function get_workbooks(string $path = 'uploads'): array {
        if (!is_dir($path))
            throw new \InvalidArgumentException("$path is not a valid directory!");

        $files = array_diff(scandir($path), ['..', '.']);

        if (empty($files))
            throw new FileNotFoundException("$path does not contain any files!");

        $this->no_of_workbooks = count($files);

        return array_map(fn ($file) => $path . DIRECTORY_SEPARATOR . $file, $files);
    }

    private function create_workbook(): void {

        $spreadsheet = new Spreadsheet();

        $spreadsheet->getProperties()
            ->setCreator("Peter")
            ->setLastModifiedBy("Peter")
            ->setTitle("Office 2007 XLSX Test Document")
            ->setSubject("Office 2007 XLSX Test Document")
            ->setDescription(
                "Test document for Office 2007 XLSX, generated using PHP classes."
            );
        // ->setKeywords("office 2007 openxml php")
        // ->setCategory("Test result file");


        $spreadsheet = $this->new_spreadsheet;
    }


    private function write_stats(string $operation, array $statistics): void {

        $stats_sheet = $this->new_spreadsheet->getActiveSheet();
        $stats_sheet->setTitle(ucwords("$operation Stats"));

        $summary = sprintf(
            "%s %s %s with a total of %d sheets",
            $operation,
            ($operation == 'converted') ? 'a' : (string)$this->no_of_workbooks,
            ($operation == 'converted') ? 'workbook' : 'workbooks',
            $this->total_sheets
        );

        // merged x workbooks with a total of x sheets
        // converted a workbook with a total of x sheets

        $stats_sheet->setCellValue('A1', $summary);

        if (empty($statistics))
            throw new \InvalidArgumentException("stats seems to empty!");

        $row = 3;

        foreach ($statistics as $stat) {
            $stats_sheet->setCellValue('A' . $row, $stat['filename']);
            $stats_sheet->setCellValue('B' . $row, $stat['sheet_count'] . ' sheets');
            $row++;
        }


        $this->save_workbook(name: "output");
    }
    private function save_workbook(string $name) {

        $writer = IOFactory::createWriter($this->new_spreadsheet, 'Xlsx');

        $output_file = OUTPUT_DIR . DIRECTORY_SEPARATOR . date('YmdHis') . "$name.xlsx";

        $writer->save($output_file);
    }
    private function copy_sheet_content(Worksheet $source_sheet, Worksheet $target_sheet): void {
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
}
