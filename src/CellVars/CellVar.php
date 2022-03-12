<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

abstract class CellVar implements CellVarInterface
{
    private $data;
    private $renderDirection;
    private $isInsertNew;
    private $callback;
    private $originCellValue;
    private $columnAndRow;
    private $shouldInsertRows = 0;
    private $shouldInsertCols = 0;

    abstract public function getCellSetter();

    protected function setData($data)
    {
        $this->data = $data;
    }

    public function getData()
    {
        return $this->data;
    }

    protected function setRenderDirection($direction)
    {
        if (!$direction instanceof RenderDirection && is_string($direction)) {
            $direction = new RenderDirection($direction);
        }
        $this->renderDirection = $direction;
    }

    public function getRenderDirection(): RenderDirection
    {
        return $this->renderDirection;
    }

    protected function setIsInsertNew(bool $isInsertNew)
    {
        $this->isInsertNew = $isInsertNew;
    }

    public function getIsInsertNew(): bool
    {
        return $this->isInsertNew;
    }

    public function setCallback($callback): void
    {
        $this->callback = $callback;
    }

    public function getCallback()
    {
        return $this->callback;
    }

    public function setOriginCellValue($originCellValue): void
    {
        $this->originCellValue = $originCellValue;
    }

    public function getOriginCellValue()
    {
        return $this->originCellValue;
    }

    public function setColumnAndRow(array $columnAndRow)
    {
        $this->columnAndRow = $columnAndRow;
    }

    public function getColumnAndRow(): array
    {
        return $this->columnAndRow;
    }

    public function setShouldInsertRowsAndCols()
    {
        $rows = $cols = 0;
        if ($this->getIsInsertNew()) {
            if ($this instanceof CellArrayVar) {
                if ($this->getRenderDirection()->isDirection(RenderDirection::DOWN)) {
                    $rows = count($this->getData()) - 1;
                }
                if ($this->getRenderDirection()->isDirection(RenderDirection::RIGHT)) {
                    $cols = count($this->getData()) - 1;
                }
            }
            //TODO 二维
        }
        $this->shouldInsertCols = $cols;
        $this->shouldInsertRows = $rows;
    }

    public function getShouldInsertRows(): int
    {
        return $this->shouldInsertRows;
    }

    public function getShouldInsertCols(): int
    {
        return $this->shouldInsertCols;
    }
}
