<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use Kaxiluo\PhpExcelTemplate\ExcelRenderContext;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellArray2DSetter extends CellArraySetter
{
    protected static function setCellVale(CellVarInterface $cellVar, Worksheet $worksheet, ExcelRenderContext $context)
    {
        $array = $cellVar->getData();
        list($col, $row) = $cellVar->getColumnAndRow();

        $startCol = $col;
        foreach ($array as $values) {
            $col = $startCol;
            foreach ($values as $value) {
                $worksheet->setCellValueByColumnAndRow($col, $row, $value);
                if (!in_array($col, $context->insertedColIndexes) && !in_array($row, $context->insertedRowIndexes)) {
                    $context->addSkipRowAndCol($row, $col);
                }
                $col++;
            }
            $row++;
        }
    }
}
