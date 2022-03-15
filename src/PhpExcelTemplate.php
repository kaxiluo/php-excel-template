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

        if ($vars) {
            static::render($worksheet, $vars);
        }

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

        $context = new ExcelRenderContext();

        $row = 1;
        while ($row <= $maxRow) {
            $col = 1;

            // 重置该行最大插入行数
            $context->iMaxInsertRows = 0;

            while ($col <= $maxCol) {
                // 跳过新列
                if (in_array($col, $context->insertedColIndexes)) {
                    $col++;
                    continue;
                }
                // 跳过非插入新行或列的用户已渲染坐标
                if ($context->hasSkipRowAndCol($row, $col)) {
                    $col++;
                    continue;
                }

                $cellValue = (string)$worksheet->getCellByColumnAndRow($col, $row)->getValue();

                // match var
                $cellVar = $varMatcher->matchCellVar($cellValue);
                if (!$cellVar) {
                    $col++;
                    continue;
                }
                $cellVar->setOriginCellValue($cellValue);
                $cellVar->setColumnAndRow([$col, $row]);

                $cellVar->getCellSetter()::render($cellVar, $worksheet, $context);

                // 扩大maxCol
                if ($context->perCellInsertedCol) {
                    $maxCol += $context->perCellInsertedCol;
                }

                $col++;
            }

            // 扩大maxRow 跳过新行
            if ($context->iMaxInsertRows) {
                $maxRow += $context->iMaxInsertRows;
                $row += $context->iMaxInsertRows;
            }

            $row++;
        }
    }
}
