<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use Kaxiluo\PhpExcelTemplate\ExcelRenderContext;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

interface CellSetterInterface
{
    public static function render(CellVarInterface $cellVar, Worksheet $worksheet, ExcelRenderContext $context);
}
