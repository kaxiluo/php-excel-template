<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use Kaxiluo\PhpExcelTemplate\CellSetter\CellArraySetter;

class CellArrayVar extends CellVar
{
    const VAR_PATTERN = '/\[([a-zA-Z-_\.\d]+)\]/';

    public function __construct(array $data, $direction = RenderDirection::DOWN, bool $isInsertNew = true)
    {
        if (!RenderDirection::isValid($direction)) {
            throw new \UnexpectedValueException('CellArrayVar Unexpected RenderDirection [' . $direction . ']');
        }
        $this->setData($data);
        $this->setRenderDirections([$direction]);
        $this->setIsInsertNew($isInsertNew);

        $this->setShouldInsertRowsAndCols();
    }

    public function setShouldInsertRowsAndCols()
    {
        if ($this->getIsInsertNew()) {
            $rows = $cols = 0;
            if ($this->hasRenderDirection(RenderDirection::DOWN)) {
                $rows = count($this->getData()) - 1;
            }
            if ($this->hasRenderDirection(RenderDirection::RIGHT)) {
                $cols = count($this->getData()) - 1;
            }
            $this->setShouldInsertRows($rows);
            $this->setShouldInsertCols($cols);
        }
    }

    public function getCellSetter(): string
    {
        return CellArraySetter::class;
    }
}
