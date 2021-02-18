<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 13:49.
 */

namespace app\office\service\filter;

class ExcelRowFilter
{
    protected $data = [];

    protected $row = 1;

    /**
     * ExcelRowFilter constructor.
     * @param  array  $data
     * @param  int  $row
     */
    public function __construct(array $data, int $row = 1)
    {
        $this->data = $data;
        $this->row = $row;
    }

    /**
     * @return array
     */
    public function getData(): array
    {
        return $this->data;
    }

    /**
     * @param  array  $data
     */
    public function setData(array $data): void
    {
        $this->data = $data;
    }


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