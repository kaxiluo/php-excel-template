<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellStringSetter;

class CellStringVar extends CellVar
{
    const VAR_PATTERN = '/\{([a-zA-Z-_\.\d]+)\}/';

    public function __construct(string $data)
    {
        $this->setData($data);
        $this->setIsInsertNew(false);
    }

    public function getCellSetter(): string
    {
        return CellStringSetter::class;
    }
}
