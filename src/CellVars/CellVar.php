<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

abstract class CellVar implements CellVarInterface
{
    private $data;
    private $renderDirections = [];
    private $isInsertNew;
    private $callback;
    private $originCellValue;
    private $columnAndRow;
    private $shouldInsertRows = 0;
    private $shouldInsertCols = 0;

    abstract public function getCellSetter();

    protected function setData($data)
    {
        if (is_array($data) && empty($data)) {
            throw new \UnexpectedValueException('Data cannot be empty when it is an array');
        }
        $this->data = $data;
    }

    public function getData()
    {
        return $this->data;
    }

    protected function setRenderDirections(array $directions)
    {
        $this->renderDirections = $directions;
    }

    public function getRenderDirections(): array
    {
        return $this->renderDirections;
    }

    public function hasRenderDirection($direction): bool
    {
        return in_array($direction, $this->renderDirections);
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

    public function setShouldInsertRows($rows): void
    {
        $this->shouldInsertRows = $rows;
    }

    public function getShouldInsertRows(): int
    {
        return $this->shouldInsertRows;
    }

    public function setShouldInsertCols($cols): void
    {
        $this->shouldInsertCols = $cols;
    }

    public function getShouldInsertCols(): int
    {
        return $this->shouldInsertCols;
    }
}
