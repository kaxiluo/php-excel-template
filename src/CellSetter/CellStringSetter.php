<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CallbackContext;
use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use Kaxiluo\PhpExcelTemplate\ExcelRenderContext;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellStringSetter implements CellSetterInterface
{
    public static function render(CellVarInterface $cellVar, Worksheet $worksheet, ExcelRenderContext $context)
    {
        $newString = preg_replace_callback($cellVar::VAR_PATTERN, function ($matches) use ($cellVar, $context) {
            if (isset($context->cellVars[$matches[1]])) {
                return (string)$context->cellVars[$matches[1]];
            } else {
                return $matches[0];
            }
        }, $cellVar->getOriginCellValue());

        list($col, $row) = $cellVar->getColumnAndRow();

        $worksheet->setCellValueByColumnAndRow($col, $row, $newString);

        if ($cellVar->getCallback()) {
            call_user_func(
                $cellVar->getCallback(),
                new CallbackContext($worksheet, $row, $col, $newString, 0, 0)
            );
        }
    }
}
