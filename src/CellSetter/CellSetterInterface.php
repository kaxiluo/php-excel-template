<?php

namespace Kaxiluo\PhpExcelTemplate\CellSetter;

use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

interface CellSetterInterface
{
    public static function render(CellVarInterface $cellVar, Worksheet $worksheet, &$context);
}
