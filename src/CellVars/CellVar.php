<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

abstract class CellVar implements CellVarInterface
{
    private $data;
    private $callback;
    private $originCellValue;
    private $columnAndRow;

    abstract public function getCellSetter();

    protected function setData($data)
    {
        $this->data = $data;
    }

    public function getData()
    {
        return $this->data;
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
}
