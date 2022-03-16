<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CallbackContext;
use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use Kaxiluo\PhpExcelTemplate\CellVars\RenderDirection;
use Kaxiluo\PhpExcelTemplate\ExcelRenderContext;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellArray2DSetter extends CellArraySetter
{
    protected static function setCellVale(CellVarInterface $cellVar, Worksheet $worksheet, ExcelRenderContext $context)
    {
        $array = $cellVar->getData();
        list($col, $row) = $cellVar->getColumnAndRow();

        $startCol = $col;
        foreach ($array as $rowKey => $values) {
            $col = $startCol;
            foreach ($values as $colKey => $value) {
                $worksheet->setCellValueByColumnAndRow($col, $row, $value);
                if ($cellVar->getCallback()) {
                    call_user_func(
                        $cellVar->getCallback(),
                        new CallbackContext($worksheet, $row, $col, $value, $rowKey, $colKey)
                    );
                }

                if (!in_array($col, $context->insertedColIndexes) && !in_array($row, $context->insertedRowIndexes)) {
                    $context->addSkipRowAndCol($row, $col);
                }
                $col++;
            }
            $row++;
        }
    }
}
