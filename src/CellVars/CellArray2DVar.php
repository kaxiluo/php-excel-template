<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellArray2DSetter;

class CellArray2DVar extends CellVar
{
    const VAR_PATTERN = '/\[\[([a-zA-Z-_\.\d]+)\]\]/';

    private $downIsInsertNew;
    private $rightIsInsertNew;

    public function __construct(array $data, bool $downIsInsertNew = true, bool $rightIsInsertNew = false)
    {
        $this->setData($data);
        $this->setRenderDirections([RenderDirection::DOWN, RenderDirection::RIGHT]);
        $this->setIsInsertNew($downIsInsertNew || $rightIsInsertNew);
        $this->downIsInsertNew = $downIsInsertNew;
        $this->rightIsInsertNew = $rightIsInsertNew;

        $this->setShouldInsertRowsAndCols();
    }

    public function setShouldInsertRowsAndCols()
    {
        if ($this->downIsInsertNew) {
            $this->setShouldInsertRows(count($this->getData()) - 1);
        }
        if ($this->rightIsInsertNew) {
            $this->setShouldInsertCols(count($this->getData()[0]) - 1);
        }
    }

    public function getCellSetter(): string
    {
        return CellArray2DSetter::class;
    }
}
