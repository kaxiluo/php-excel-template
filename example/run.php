<?php

use Kaxiluo\PhpExcelTemplate\CellVars\CallbackContext;
use Kaxiluo\PhpExcelTemplate\CellVars\CellArray2DVar;
use Kaxiluo\PhpExcelTemplate\CellVars\CellArrayVar;
use Kaxiluo\PhpExcelTemplate\CellVars\RenderDirection;
use Kaxiluo\PhpExcelTemplate\PhpExcelTemplate;

require_once '../vendor/autoload.php';

$data = [
    ['个人任务执行情况', '完成X系统开发', '40', '35', '任务延期一次'],
    ['技术贡献', '主持基础设施项目建设', '30', '30', ''],
    ['素质与能力', '认真负责，积极工作', '20', '20', ''],
    ['防疫工作', '严格执行防疫措施，按要求填写防疫信息统计表', '10', '10', ''],
];
$items = new CellArray2DVar(
    $data,
    true,
    false,
    function (CallbackContext $context) use ($data) {
        if ($context->getLoopColKey() === 3) {
            if ($context->getValue() < $data[$context->getLoopRowKey()][2]) {
                $context->getStyle()->getFont()->getColor()->setARGB('FFFF0000');
            }
        }
    }
);
$vars = [
    'username' => 'lyy',
    'department' => 'IT中心',
    'dateRange' => '2022-03-01 - 2022-03-31',
    'leader' => 'Tim',
    'items' => $items,
    'totalScore' => '=SUM(D5:D' . (4 + count($data)) . ')',
    'x' => new CellArrayVar(['Peace', 'and', 'Love'], RenderDirection::DOWN, false),
];

PhpExcelTemplate::save('./example-kpi.xlsx', './example-kpi-output.xlsx', $vars);
