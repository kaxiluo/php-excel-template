<?php

namespace Tests;

use Kaxiluo\PhpExcelTemplate\CellVars\CellArrayVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellStringVar;
use Kaxiluo\PhpExcelTemplate\CellVars\RenderDirection;
use Kaxiluo\PhpExcelTemplate\PhpExcelTemplate;

class SampleExcelTemplateTest extends TestCase
{
    public function testSimple()
    {
        $template = __DIR__ . '/template-excel/sample-test.xlsx';
        $outputFile = $this->getOutputFile($template);
        $vars = [
            'A1-A2' => ['a1', 'a2'],
            'B1-B2' => new CellArrayVar(['b1', 'b2']),
            'C1-D1' => new CellArrayVar(['c1', 'd1'], RenderDirection::RIGHT),
            'x' => new CellArrayVar(['a4', 'b4', 'z'], RenderDirection::RIGHT, true),
            'xx' => new CellArrayVar(['g1', 'g2'], RenderDirection::RIGHT, true),
        ];

        PhpExcelTemplate::save($template, $outputFile, $vars);
    }

    public function testStringVar()
    {
        $template = __DIR__ . '/template-excel/sample-string.xlsx';
        $outputFile = $this->getOutputFile($template);
        $vars = [
            'A1' => 'aa',
            'B1' => new CellStringVar('bb'),
            'C1' => new CellArrayVar(['cc']),
            'A2' => new CellStringVar('a2'),
            'C2' => new CellStringVar('c2'),
        ];

        PhpExcelTemplate::save($template, $outputFile, $vars);

        $this->assertExcelCellValue($outputFile, [
            'A1' => 'aa',
            'B1' => 'bb',
            'C1' => '{C1}',//类型错误，不渲染
            'D1' => '{D1}',//未定义该变量，不渲染
            'A2' => 'hi a2',//包含
            'B2' => 'aa',//重复使用
            'C2' => 'hello c2 -',//包含
        ]);
    }

    public function testArrayVar()
    {
        $template = __DIR__ . '/template-excel/sample-array.xlsx';
        $outputFile = $this->getOutputFile($template);
        $vars = [
            'A1-A2' => ['a1', 'a2'],
            'B1-B2' => new CellArrayVar(['b1', 'b2']),
            'C1-D1' => new CellArrayVar(['c1', 'd1'], RenderDirection::RIGHT),
            'C2-E2' => new CellArrayVar(['c2', 'd2', 'e2'], RenderDirection::RIGHT),
            'A4-B4' => new CellArrayVar(['a4', 'b4',], RenderDirection::RIGHT, false),
            'G1-G3' => new CellArrayVar(['g1', 'g2', 'g3'], RenderDirection::DOWN, false),

        ];

        PhpExcelTemplate::save($template, $outputFile, $vars);

        $this->assertExcelCellValue($outputFile, [
            // 定义为数组 向下插入新行
            'A1' => 'a1',
            'A2' => 'a2',
            // 定义为对象 向下插入新行 同行不额外插入
            'B1' => 'b1',
            'B2' => 'b2',
            // 向右插入新列
            'C1' => 'c1',
            'D1' => 'd1',
            // 向右插入新列 同列不额外插入 受新行影响行2偏移到3
            'C3' => 'c2',
            'D3' => 'd2',
            'E3' => 'e2',
            // 向右 不插入新行
            'A4' => 'a4',
            'B4' => 'b4',
            // 向下 不插入新行 受新列影响G偏移到I
            'I1' => 'g1',
            'I3' => 'g3',
            'I5' => 'I am origin g4',
        ]);
    }
}
