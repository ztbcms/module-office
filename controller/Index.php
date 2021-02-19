<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 09:20.
 */

namespace app\office\controller;

use app\common\controller\AdminController;
use app\office\service\ExcelService;
use app\office\service\filter\common\DateFilter;
use app\office\service\filter\demo\GoodsFilter;
use app\office\service\filter\ExcelRowFilter;
use app\office\service\Import\DemoImportTemplate;
use think\facade\Filesystem;
use think\facade\View;

class Index extends AdminController
{
    function index()
    {
        return View::fetch('index');
    }

    /**
     * xls 导入
     * @return array|string
     */
    function importXls()
    {
        $filePath = request()->param('filepath');
        try {
            $excelImportService = new DemoImportTemplate(Filesystem::disk('ztbcms')->path('/').$filePath);
            return self::makeJsonReturn(true, $excelImportService->getData());
        } catch (\Throwable $exception) {
            return self::makeJsonReturn(false, [], $exception->getMessage());
        }
    }

    /**
     * xls导出
     * @return array|string
     */
    function exportXls()
    {
        $fileName = '';
        $fileHeader = ["订单号", "订单总价", "订单商品", "下单时间"];
        $exportData = [
            new ExcelRowFilter(
                [
                    'a' => 1,
                    'b' => 1.1,
                    'c' => new GoodsFilter(['hello', 'world', 'excel'], 3),
                    'd' => new DateFilter(time())
                ], 3
            ),
            [
                'a' => 2,
                'b' => 2.1,
                'c' => ['hello', 'world', 'excel'],
                'd' => new DateFilter(time())
            ], [
                'a' => 3,
                'b' => 3.1,
                'c' => "hello 3",
                'd' => date('Y-m-d H:i:s', time())
            ]
        ];
        $ex = new ExcelService($fileName, $fileHeader, $exportData);
        try {
            return self::makeJsonReturn(true, $ex->save());
        } catch (\Throwable $exception) {
            return $exception->getMessage();
        }
    }
}