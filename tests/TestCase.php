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
            $actualValue = $worksheet->getCell($coordinate)->getFormattedValue();
            $this->assertEquals($value, $actualValue);
        }
    }

    protected function assertExcelCellStyle($excelFile, array $coordinateToValue)
    {
        $spreadsheet = IOFactory::load($excelFile);
        $worksheet = $spreadsheet->getActiveSheet();
        foreach ($coordinateToValue as $coordinate => $value) {
            $actual = $worksheet->getStyle($coordinate)->exportArray();

            $this->assertArraySubset($value, $actual);
        }
    }
}
