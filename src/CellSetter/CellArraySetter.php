<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use Kaxiluo\PhpExcelTemplate\CellVars\RenderDirection;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CellArraySetter implements CellSetterInterface
{
    public static function render(CellVarInterface $cellVar, Worksheet $worksheet, &$context)
    {
        list($col, $row) = $cellVar->getColumnAndRow();
        $array = $cellVar->getData();
        $direction = $cellVar->getRenderDirection();

        $context['perCellInsertedCol'] = 0;

        // 插入行
        if ($shouldInsertRows = $cellVar->getShouldInsertRows()) {
            if ($context['iMaxInsertRows'] < $shouldInsertRows) {
                $worksheet->insertNewRowBefore(
                    $row + $context['iMaxInsertRows'] + 1,
                    $shouldInsertRows - $context['iMaxInsertRows']
                );
                // 更新该行最大插入新行数
                $context['iMaxInsertRows'] = $shouldInsertRows;
            }
        }

        // 插入列
        if ($shouldInsertCols = $cellVar->getShouldInsertCols()) {
            $insertCol = [];
            if (isset($context['colToMaxInsertCols'][$col])) {
                // 当前列所需插入的列数 大于 当列之前已经插入的列数 插入缺少的列数
                if ($shouldInsertCols > $context['colToMaxInsertCols'][$col]) {
                    $insertCol = [
                        $col + $context['colToMaxInsertCols'][$col],
                        $shouldInsertCols - $context['colToMaxInsertCols'][$col]
                    ];
                    $context['colToMaxInsertCols'][$col] = $shouldInsertCols;
                }
            } else {
                $insertCol = [$col, $shouldInsertCols];
                $context['colToMaxInsertCols'][$col] = $shouldInsertCols;
            }

            if ($insertCol) {
                $worksheet->insertNewColumnBeforeByIndex($insertCol[0] + 1, $insertCol[1]);
                for ($i = $insertCol[0]; $i < ($col + $shouldInsertCols); $i++) {
                    $context['insertedColIndexes'][] = $i + 1;
                }
                $context['perCellInsertedCol'] = $insertCol[1];
            }
        }

        foreach ($array as $value) {
            $worksheet->setCellValueByColumnAndRow($col, $row, $value);

            if ($direction->isDirection(RenderDirection::DOWN)) {
                $row++;
            }
            if ($direction->isDirection(RenderDirection::RIGHT)) {
                $col++;
            }
        }
    }
}
