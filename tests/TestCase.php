<?php

namespace Tests;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PHPUnit\Framework\TestCase as BaseTestCase;

abstract class TestCase extends BaseTestCase
{
    protected function getOutputFile($templateFile)
    {
        return str_replace('.xlsx', '-output.xlsx', $templateFile);
    }

    protected function assertExcelCellValue($excelFile, array $coordinateToValue)
    {
        $spreadsheet = IOFactory::load($excelFile);
        $worksheet = $spreadsheet->getActiveSheet();
        foreach ($coordinateToValue as $coordinate => $value) {
            $actualValue = (string)$worksheet->getCell($coordinate)->getValue();
            $this->assertEquals($value, $actualValue);
        }
    }
}