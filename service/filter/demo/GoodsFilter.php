<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 14:40.
 */

namespace app\office\service\filter\demo;


use app\office\service\filter\ExcelValueFilter;

class GoodsFilter extends ExcelValueFilter
{

    public function getFilterValue(): array
    {
        return $this->value;
    }
}