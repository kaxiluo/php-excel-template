<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellArraySetter;

class CellArrayVar extends CellVar
{
    const VAR_PATTERN = '/\[([a-zA-Z_\.\d]+)\]/';

    private $direction;

    private $isInsertNew;

    public function __construct(array $data, $direction = RenderDirection::DOWN, bool $isInsertNew = true)
    {
        $this->setData($data);
        $this->direction = $direction;
        $this->isInsertNew = $isInsertNew;
    }

    public function getDirection()
    {
        return $this->direction;
    }

    public function getIsInsertNew(): bool
    {
        return $this->isInsertNew;
    }

    public function getCellSetter(): string
    {
        return CellArraySetter::class;
    }
}
