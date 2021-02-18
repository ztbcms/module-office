<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:54.
 */

namespace app\office\service\filter\common;


use app\office\service\filter\ExcelValueFilter;

class DateFilter extends ExcelValueFilter
{
    public function getFilterValue(): string
    {
        if (is_numeric($this->value)) {
            return date("Y-m-d H:i:s", $this->value);
        } else {
            return date("Y-m-d H:i:s", 0);
        }
    }
}