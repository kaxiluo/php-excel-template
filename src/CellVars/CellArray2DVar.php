<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellArray2DSetter;

class CellArray2DVar extends CellVar
{
    const VAR_PATTERN = '/\[\[([a-zA-Z-_\.\d]+)\]\]/';

    public function __construct(array $data, bool $isInsertNew = true)
    {
        $this->setData($data);
        $this->setIsInsertNew($isInsertNew);
    }

    public function getCellSetter(): string
    {
        return CellArray2DSetter::class;
    }
}
