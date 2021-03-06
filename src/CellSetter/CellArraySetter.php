<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CallbackContext;
use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use Kaxiluo\PhpExcelTemplate\CellVars\RenderDirection;
use Kaxiluo\PhpExcelTemplate\ExcelRenderContext;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellArraySetter implements CellSetterInterface
{
    public static function render(CellVarInterface $cellVar, Worksheet $worksheet, ExcelRenderContext $context)
    {
        list($col, $row) = $cellVar->getColumnAndRow();

        $context->perCellInsertedCol = 0;

        // 插入行
        if ($shouldInsertRows = $cellVar->getShouldInsertRows()) {
            if ($context->iMaxInsertRows < $shouldInsertRows) {
                $insertRow = [$row + $context->iMaxInsertRows, $shouldInsertRows - $context->iMaxInsertRows];
                $worksheet->insertNewRowBefore($insertRow[0] + 1, $insertRow[1]);
                // 更新该行最大插入新行数
                $context->iMaxInsertRows = $shouldInsertRows;
                // 记录插入的列
                $context->insertedRowIndexes += range($insertRow[0] + 1, $row + $shouldInsertRows);
                // 动态更新skip row
                $context->dynamicUpdateSkipRow($insertRow[0] + 1, $insertRow[1]);
            }
        }

        // 插入列
        if ($shouldInsertCols = $cellVar->getShouldInsertCols()) {
            $insertCol = [];
            if (isset($context->colToMaxInsertCols[$col])) {
                // 当前列所需插入的列数 大于 当列之前已经插入的列数 插入缺少的列数
                if ($shouldInsertCols > $context->colToMaxInsertCols[$col]) {
                    $insertCol = [
                        $col + $context->colToMaxInsertCols[$col],
                        $shouldInsertCols - $context->colToMaxInsertCols[$col]
                    ];
                    $context->colToMaxInsertCols[$col] = $shouldInsertCols;
                }
            } else {
                $insertCol = [$col, $shouldInsertCols];
                $context->colToMaxInsertCols[$col] = $shouldInsertCols;
            }

            if ($insertCol) {
                $worksheet->insertNewColumnBeforeByIndex($insertCol[0] + 1, $insertCol[1]);
                // 记录插入的行
                $context->insertedColIndexes += range($insertCol[0] + 1, $col + $shouldInsertCols);
                $context->perCellInsertedCol = $insertCol[1];
                // 动态更新skip col
                $context->dynamicUpdateSkipCol($insertCol[0] + 1, $insertCol[1]);
            }
        }

        static::setCellVale($cellVar, $worksheet, $context);
    }

    protected static function setCellVale(CellVarInterface $cellVar, Worksheet $worksheet, ExcelRenderContext $context)
    {
        $array = $cellVar->getData();
        list($col, $row) = $cellVar->getColumnAndRow();

        foreach ($array as $key => $value) {
            $worksheet->setCellValueByColumnAndRow($col, $row, $value);
            if ($cellVar->getCallback()) {
                $loopRowKey = $loopColKey = 0;
                if ($cellVar->hasRenderDirection(RenderDirection::DOWN)) {
                    $loopRowKey = $key;
                }
                if ($cellVar->hasRenderDirection(RenderDirection::RIGHT)) {
                    $loopColKey = $key;
                }
                call_user_func(
                    $cellVar->getCallback(),
                    new CallbackContext($worksheet, $row, $col, $value, $loopRowKey, $loopColKey)
                );
            }

            if ($key !== 0 && !in_array($col, $context->insertedColIndexes)
                && !in_array($row, $context->insertedRowIndexes)) {
                $context->addSkipRowAndCol($row, $col);
            }

            if ($cellVar->hasRenderDirection(RenderDirection::DOWN)) {
                $row++;
            }
            if ($cellVar->hasRenderDirection(RenderDirection::RIGHT)) {
                $col++;
            }
        }
    }
}
