<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellStringSetter implements CellSetterInterface
{
    public static function render(CellVarInterface $cellVar, Worksheet $worksheet)
    {
        $newString = preg_replace_callback($cellVar::VAR_PATTERN, function () use ($cellVar) {
            return $cellVar->getData();
        }, $cellVar->getOriginCellValue());
        list($col, $row) = $cellVar->getColumnAndRow();
        $worksheet->setCellValueByColumnAndRow($col, $row, $newString);
    }
}