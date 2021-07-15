<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:54.
 */

namespace app\office\service\filter\common;


use app\office\service\filter\ExcelValueFilter;

class StringToNumber extends ExcelValueFilter
{
    public function getFilterValue(): float
    {
        if (is_string($this->value)) {
            $res = floatval((float) $this->value);
            return is_numeric($res) ? $res : 0;
        } else {
            return 0;
        }
    }
}