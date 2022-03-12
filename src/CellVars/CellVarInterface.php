<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellSetterInterface;

interface CellVarInterface
{
    public function getData();

    public function getIsInsertNew(): bool;

    public function getRenderDirection(): RenderDirection;

    public function setCallback($callback);

    public function getCallback();

    public function setOriginCellValue($originCellValue);

    public function getOriginCellValue();

    public function setColumnAndRow(array $columnAndRow);

    public function getColumnAndRow(): array;

    public function getShouldInsertRows(): int;

    public function getShouldInsertCols(): int;

    /**
     * @return CellSetterInterface
     */
    public function getCellSetter();
}
