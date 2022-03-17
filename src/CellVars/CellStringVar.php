<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellStringSetter;

class CellStringVar extends CellVar
{
    const VAR_PATTERN = '/\{([a-zA-Z-_\.\d]+)\}/';

    public function __construct(string $data, ?callable $callback = null)
    {
        $this->setData($data);
        $this->setIsInsertNew(false);
        $this->setCallback($callback);
    }

    public function getCellSetter(): string
    {
        return CellStringSetter::class;
    }

    public function __toString() : string
    {
        return $this->getData();
    }
}
