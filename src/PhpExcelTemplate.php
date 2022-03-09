<?php

namespace Kaxiluo\PhpExcelTemplate;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class PhpExcelTemplate
{
    public static function save($templateFile, $outputFile, array $vars)
    {
        $spreadsheet = IOFactory::load($templateFile);
        $worksheet = $spreadsheet->getActiveSheet();

        static::render($worksheet, $vars);

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($outputFile);
        return $outputFile;
    }

    protected static function render(Worksheet $worksheet, array $vars)
    {
        $maxRow = $worksheet->getHighestRow();
        $highestCol = $worksheet->getHighestColumn();
        $maxCol = Coordinate::columnIndexFromString($highestCol);

        $varMatcher = new CellVarMatcher($vars);

        $row = 1;
        while ($row <= $maxRow) {
            for ($col = 1; $col <= $maxCol; $col++) {
                $cellValue = (string)$worksheet->getCellByColumnAndRow($col, $row)->getValue();
                $cellVar = $varMatcher->matchCellVar($cellValue);
                if (empty($cellVar)) {
                    continue;
                }
                $cellVar->setOriginCellValue($cellValue);
                $cellVar->setColumnAndRow([$col, $row]);
                // render
                $cellVar->getCellSetter()::render($cellVar, $worksheet);
            }
            $row = $row + 1;
        }
    }
}
