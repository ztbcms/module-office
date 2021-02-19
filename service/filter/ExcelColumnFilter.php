<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:49.
 */

namespace app\office\service\filter;


class ExcelColumnFilter
{
    public $key = '';
    public $filter = null;

    /**
     * ExcelColumnFilter constructor.
     * @param  string  $key
     * @param  null  $filter
     */
    public function __construct(string $key, $filter = null)
    {
        $this->key = $key;
        $this->filter = $filter;
    }

}