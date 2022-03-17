<?php

namespace Kaxiluo\PhpExcelTemplate;

class ExcelRenderContext
{
    public $skipRowAndCol = [];// 非新插入的、被渲染的行-列
    public $iMaxInsertRows = 0;// 单行最大新行数
    public $insertedRowIndexes = [];// 插入的新行索引
    public $colToMaxInsertCols = [];// 列 -> 最大插入列数
    public $insertedColIndexes = [];// 插入的新列索引
    public $perCellInsertedCol = 0;// 每个单元格渲染后 真正插入的列数
    public $cellVars = [];// 所有变量

    public function addSkipRowAndCol($row, $col)
    {
        $this->skipRowAndCol[] = [$row, $col];
    }

    public function hasSkipRowAndCol($row, $col): bool
    {
        foreach ($this->skipRowAndCol as $item) {
            if ($item == [$row, $col]) {
                return true;
            }
        }
        return false;
    }

    public function dynamicUpdateSkipRow($beforeRow, $numberOfRows)
    {
        foreach ($this->skipRowAndCol as $key => $item) {
            if ($item[0] >= $beforeRow) {
                $this->skipRowAndCol[$key][0] += $numberOfRows;
            }
        }
    }

    public function dynamicUpdateSkipCol($beforeCol, $numberOfCols)
    {
        foreach ($this->skipRowAndCol as $key => $item) {
            if ($item[1] >= $beforeCol) {
                $this->skipRowAndCol[$key][1] += $numberOfCols;
            }
        }
    }
}
