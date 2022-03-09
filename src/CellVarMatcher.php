<?php

namespace Kaxiluo\PhpExcelTemplate;

use Kaxiluo\PhpExcelTemplate\CellVars\CellArray2DVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellArrayVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellStringVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellVarInterface;

class CellVarMatcher
{
    private $cellVars = [];

    public function __construct(array $cellVars)
    {
        $this->setCellVars($cellVars);
    }

    /**
     * @return CellVarInterface[]|CellVarInterface|null
     */
    public function getCellVars($key = '')
    {
        if ($key) {
            return $this->cellVars[$key] ?? null;
        } else {
            return $this->cellVars;
        }
    }

    private function setCellVars(array $cellVars): void
    {
        foreach ($cellVars as $key => $cellVar) {
            if ($cellVar instanceof CellVarInterface) {
                continue;
            }
            if (is_string($cellVar)) {
                $cellVars[$key] = new CellStringVar($cellVar);
            } elseif (is_array($cellVar)) {
                $cellVar = array_values($cellVar);
                if (count($cellVar) == count($cellVar, 1)) {
                    $cellVars[$key] = new CellArrayVar($cellVar);
                } else {
                    $cellVars[$key] = new CellArray2DVar($cellVar);
                }
            } else {
                throw new \UnexpectedValueException('Unexpected Cell Var [' . $key . ']');
            }
        }
        $this->cellVars = $cellVars;
    }

    /**
     * @return CellVar[]
     */
    protected function getCellVarClasses(): array
    {
        return [CellStringVar::class, CellArrayVar::class, CellArray2DVar::class];
    }

    /**
     * @param $cellValue
     * @return CellVarInterface|null
     */
    public function matchCellVar($cellValue): ?CellVarInterface
    {
        $cellVar = null;
        foreach ($this->getCellVarClasses() as $varClass) {
            $varName = static::matchVarName($varClass, $cellValue);
            if (empty($varName)) {
                continue;
            }
            $cellVar = $this->getCellVars($varName);
            if ($cellVar instanceof $varClass) {
                break;
            }
        }
        return $cellVar;
    }

    private static function matchVarName($varClass, $cellValue): string
    {
        if (preg_match($varClass::VAR_PATTERN, $cellValue, $matches)) {
            return $matches[1];
        }
        return '';
    }
}
