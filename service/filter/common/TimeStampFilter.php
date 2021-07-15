<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:54.
 */

namespace app\office\service\filter\common;


use app\office\service\filter\ExcelValueFilter;

class TimeStampFilter extends ExcelValueFilter
{
    public function getFilterValue(): string
    {
        if (is_string($this->value)) {
            $date = preg_replace(['/年|月/', '/日/'], ['-', ''], $this->value);
            return strtotime($date);
        } else {
            return 0;
        }
    }
}