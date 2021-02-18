<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:49.
 */

namespace app\office\service\filter;


abstract class ExcelValueFilter
{
    protected $value;

    protected $row;

    /**
     * ExcelFilter constructor.
     * @param $value
     * @param $row
     */
    public function __construct($value, $row = 1)
    {
        $this->value = $value;
        $this->row = $row;
    }

    /**
     * @return mixed
     */
    abstract public function getFilterValue();

    /**
     * @return int|mixed
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * @param  int|mixed  $row
     */
    public function setRow($row): void
    {
        $this->row = $row;
    }

}