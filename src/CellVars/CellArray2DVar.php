<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellArray2DSetter;

class CellArray2DVar extends CellVar
{
    const VAR_PATTERN = '/\[\[([a-zA-Z_\.\d]+)\]\]/';

    private $isInsertNew;

    public function __construct(array $data, bool $isInsertNew = true)
    {
        $this->setData($data);
        $this->isInsertNew = $isInsertNew;
    }

    public function getIsInsertNew(): bool
    {
        return $this->isInsertNew;
    }

    public function getCellSetter(): string
    {
        return CellArray2DSetter::class;
    }
}
