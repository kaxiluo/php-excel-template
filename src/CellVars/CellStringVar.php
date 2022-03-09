<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellSetterInterface;
use Kaxiluo\PhpExcelTemplate\CellSetter\CellStringSetter;

class CellStringVar extends CellVar
{
    const VAR_PATTERN = '/\{([a-zA-Z_\.\d]+)\}/';

    public function __construct(string $data)
    {
        $this->setData($data);
    }

    public function getCellSetter(): string
    {
        return CellStringSetter::class;
    }
}
