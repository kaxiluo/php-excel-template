<?php

namespace Tests;

use Kaxiluo\PhpExcelTemplate\CellVars\CellArrayVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellStringVar;
use Kaxiluo\PhpExcelTemplate\PhpExcelTemplate;

class SampleExcelTemplateTest extends ExcelTemplateTest
{
    public function testStringVar()
    {
        $template = __DIR__ . '/template-excel/sample-string.xlsx';
        $outputFile = $this->getOutputFile($template);
        $vars = [
            'A1' => 'aa',
            'B1' => new CellStringVar('bb'),
            'C1' => new CellArrayVar(['cc']),
            'A2' => new CellStringVar('a2a2'),
        ];

        PhpExcelTemplate::save($template, $outputFile, $vars);

        $this->assertCellValue($outputFile, [
            'A1' => 'aa',
            'B1' => 'bb',
            'C1' => '{C1}',
            'D1' => '{D1}',
            'A2' => 'hi a2a2',
        ]);
    }
}
