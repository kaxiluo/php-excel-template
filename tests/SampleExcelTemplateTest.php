<?php

namespace Tests;

use Kaxiluo\PhpExcelTemplate\CellVars\CellArray2DVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellArrayVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellStringVar;
use Kaxiluo\PhpExcelTemplate\CellVars\RenderDirection;
use Kaxiluo\PhpExcelTemplate\PhpExcelTemplate;

class SampleExcelTemplateTest extends TestCase
{
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
            'A1' => ['a1', 'a2'],
            'B1' => new CellArrayVar(['b1', 'b2']),
            'C1' => new CellArrayVar(['c1', 'd1'], RenderDirection::RIGHT),
            'C2' => new CellArrayVar(['c2', 'd2', 'e2'], RenderDirection::RIGHT),
            'A3' => new CellArrayVar(['a4', 'b4', '[A1]'], RenderDirection::RIGHT, false),
            'G1' => new CellArrayVar(['i1', 'i2', '{i3}'], RenderDirection::DOWN, false),
            'i3' => 'i3 var',
            'x' => 'x var',
            'A8' => new CellArrayVar(['a9', 'a10', '{a11}'], RenderDirection::DOWN, false),
            'a11' => 'a11 var',
            'B9' => new CellArrayVar(['x', 'xx']),
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
            'C4' => '[A1]',// 已跳过，不会渲染用户变量
            // 向下 不插入新行 受新列影响G偏移到I
            'I1' => 'i1',
            'I2' => 'i2',
            'I3' => '{i3}',// 已跳过，不会渲染用户变量 列扩展原skip坐标自动更新(G3->H3->I3)
            'H3' => 'x var',// 列扩展原skip坐标被释放
            'I5' => 'move to i5',
            // 行扩展原skip坐标自动更新
            'A9' => 'a9',
            'A12' => '{a11}',//受到B9影响 向下移动一行
        ]);
    }

    public function testArray2DVar()
    {
        $template = __DIR__ . '/template-excel/sample-array-2d.xlsx';
        $outputFile = $this->getOutputFile($template);
        $vars = [
            'A1' => [['a1', 'b1'], ['a2', 'b2']],
            'C1' => new CellArray2DVar([['c1', 'd1'], ['c2', 'd2']]),
            'A3' => new CellArray2DVar([['a5', '{b5}'], ['a6', 'b6']], false),
            '{b5}' => 'b5 var',
            'G1' => new CellArray2DVar(
                [['g1', 'h1'], ['g2', 'h2'], ['g3', 'h3']],
                true,
                true
            ),
        ];

        PhpExcelTemplate::save($template, $outputFile, $vars);

        $this->assertExcelCellValue($outputFile, [
            // group 1
            'A1' => 'a1',
            'B1' => 'b1',
            'A2' => 'a2',
            'B2' => 'b2',
            // group 2
            'C1' => 'c1',
            'D1' => 'd1',
            'C2' => 'c2',
            'D2' => 'd2',
            // group 3
            'A4' => 'move to a4',
            'A5' => 'a5',
            'B5' => '{b5}',
            'A6' => 'a6',
            'B6' => 'b6',
            // group 4
            'G1' => 'g1',
            'H2' => 'h2',
            'G3' => 'g3',
            'I1' => 'move to i1',
        ]);
    }

//    public function testSimpleMix()
//    {
//        $template = __DIR__ . '/template-excel/sample-mix.xlsx';
//        $outputFile = $this->getOutputFile($template);
//        $vars = [
//            'z' => 'ss',
//            'a' => new CellArrayVar(['a', 'b', 'c', '[a]'], RenderDirection::RIGHT, false),
//        ];
//
//        PhpExcelTemplate::save($template, $outputFile, $vars);
//        $this->assertExcelCellValue($outputFile, []);
//    }
}
