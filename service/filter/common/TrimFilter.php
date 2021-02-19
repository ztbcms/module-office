<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:54.
 */

namespace app\office\service\filter\common;


use app\office\service\filter\ExcelValueFilter;

class TrimFilter extends ExcelValueFilter
{
    public function getFilterValue(): string
    {
        $value = trim($this->value);
        return str_replace(["/r/n", "/r", "/n"], "", $value);
    }
}