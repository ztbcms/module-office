<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/19
 * Time: 09:39.
 */

namespace app\office\service\Import;


use app\office\service\ExcelImportService;
use app\office\service\filter\common\TrimFilter;
use app\office\service\filter\ExcelColumnFilter;

class DemoImportTemplate extends ExcelImportService
{
    protected $startRow = 4;
    protected $columnCount = 5;

    public function __construct(string $filePath)
    {
        parent::__construct($filePath);
        $this->keys = [
            'id', 'company_name', 'company_description', 'work_name',
            new ExcelColumnFilter('work_level', TrimFilter::class)
        ];
    }
}