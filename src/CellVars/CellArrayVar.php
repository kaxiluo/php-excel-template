<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellArraySetter;

class CellArrayVar extends CellVar
{
    const VAR_PATTERN = '/\[([a-zA-Z-_\.\d]+)\]/';

    public function __construct(array $data, $direction = RenderDirection::DOWN, bool $isInsertNew = true)
    {
        $this->setData($data);
        $this->setRenderDirection($direction);
        $this->setIsInsertNew($isInsertNew);

        $this->setShouldInsertRowsAndCols();
    }

    public function getCellSetter(): string
    {
        return CellArraySetter::class;
    }
}
