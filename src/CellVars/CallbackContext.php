<?php

namespace Kaxiluo\PhpExcelTemplate\CellVars;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class CallbackContext
{
    private $worksheet;
    private $cellRow;
    private $cellColIndex;
    private $loopRowKey;
    private $loopColKey;
    private $value;

    public function __construct(Worksheet $worksheet, int $cellRow, int $cellColIndex, $value, $loopRowKey, $loopColKey)
    {
        $this->worksheet = $worksheet;
        $this->cellRow = $cellRow;
        $this->cellColIndex = $cellColIndex;
        $this->value = $value;
        $this->loopRowKey = $loopRowKey;
        $this->loopColKey = $loopColKey;
    }

    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
    }

    /** Numeric row coordinate of the cell
     * @return int
     */
    public function getCellRow(): int
    {
        return $this->cellRow;
    }

    /** Numeric column coordinate of the cell
     * @return int
     */
    public function getCellColIndex(): int
    {
        return $this->cellColIndex;
    }

    /** Value of the current cell
     * @return mixed
     */
    public function getValue()
    {
        return $this->value;
    }

    public function getStyle(): \PhpOffice\PhpSpreadsheet\Style\Style
    {
        return $this->getWorksheet()->getStyleByColumnAndRow($this->getCellColIndex(), $this->getCellRow());
    }

    /** Row key of loop data
     * @return mixed
     */
    public function getLoopRowKey()
    {
        return $this->loopRowKey;
    }

    /** Col key of loop data
     * @return mixed
     */
    public function getLoopColKey()
    {
        return $this->loopColKey;
    }
}
