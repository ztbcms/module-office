<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 16:24.
 */

declare(strict_types=1);

namespace app\office\service;


use app\common\service\BaseService;
use app\office\model\ImportRecord;
use app\office\service\filter\ExcelColumnFilter;
use app\office\service\filter\ExcelValueFilter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Exception;
use Throwable;

class ExcelImportService extends BaseService
{

    protected $startRow = 0;
    protected $columnCount = 0;
    protected $data = [];
    protected $keys = [];

    protected $spreadsheet;
    protected $sheet;
    protected $filePath;


    /**
     * ExcelImportService constructor.
     * @param  string  $filePath
     */
    public function __construct(string $filePath)
    {
        $this->spreadsheet = IOFactory::load($filePath);
        $this->filePath = $filePath;
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function importRecord($callback)
    {
        $status = $callback($this->getData());
        $data_md5 = md5(json_encode($this->getData()));
        $file_path = $this->filePath;
        return ImportRecord::create([
            'file_path' => $file_path,
            'data_md5'  => $data_md5,
            'status'    => $status
        ]);
    }

    /**
     * @return array
     * @throws Throwable
     */
    public function getData(): array
    {
        if (empty($this->data)) {
            return $this->importData();
        }
        return $this->data;
    }

    /**
     * 导入数据处理
     * @throws Throwable
     */
    private function importData(): array
    {
        $cells = $this->sheet->getCellCollection();
        $currentRow = $cells->getCurrentRow();
        $currentColumn = $this->col2Int($cells->getCurrentColumn()) + 1;
        if ($currentRow < $this->startRow) {
            //如果规定行数大于读取的行数，则返回错误
            $this->setError('gets the number of rows less than the start-row');
            throw new Exception($this->getError());
        }
        if ($this->columnCount > $currentColumn) {
            //如果规定列数大于读取的列数，则返回错误
            $this->setError('need column not enough');
            throw new Exception($this->getError());
        }
        if ($this->columnCount != 0) {
            $currentColumn = $this->columnCount;
        }
        $data = [];
        for ($i = $this->startRow; $i <= $currentRow; $i++) {
            $itemData = [];
            for ($j = 0; $j < $currentColumn; $j++) {
                $cellKey = $this->getChar($j).$i;
                $mergeRange = $this->sheet->getCell($cellKey)->getMergeRange();
                if ($mergeRange !== false) {
                    //如果是合并的单元格，默认使用合并第一个单元的数据
                    $cellKey = explode(':', $mergeRange)[0] ?? $cellKey;
                }
                $value = $this->sheet->getCell($cellKey)->getFormattedValue();
                if (is_array($value)) {
                    $value = json_encode($value);
                }
                if (!empty($this->keys[$j])) {
                    if ($this->keys[$j] instanceof ExcelColumnFilter) {
                        $key = $this->keys[$j]->key;
                        //如果有过滤器，则使用过滤器
                        if ($this->keys[$j]->filter && class_exists((string) $this->keys[$j]->filter)) {
                            $filterClass = $this->keys[$j]->filter;
                            //获取过滤器的值
                            $filter = new $filterClass($value);
                            if ($filter instanceof ExcelValueFilter) {
                                $value = $filter->getFilterValue();
                            }
                        }
                        $itemData[$key] = $value;
                    } else {
                        $itemData[$this->keys[$j]] = $value;
                    }
                } else {
                    $itemData[] = $value;
                }
            }
            $data[] = $itemData;
        }
        $this->data = $data;
        return $data;
    }

    /**
     * 获取列数，获取A、B、C、AA
     * @param $i
     * @return string
     */
    private function getChar($i): string
    {
        $y = ($i / 26);
        if ($y >= 1) {
            $y = intval($y);
            return chr($y + 64).chr($i - $y * 26 + 65);
        } else {
            return chr($i + 65);
        }
    }

    private function col2Int(string $str): int
    {
        $num = 0;
        $strArr = str_split($str, 1);
        $length = count($strArr);
        foreach ($strArr as $k => $v) {
            $num += ((ord($v) - ord('A') + 1) * pow(26, $length - $k - 1));
        }
        return (int) $num - 1;
    }


    /**
     * @param  array  $data
     */
    public function setData(array $data): void
    {
        $this->data = $data;
    }


    /**
     * @return array
     */
    public function getKeys(): array
    {
        return $this->keys;
    }

    /**
     * @param  array  $keys
     */
    public function setKeys(array $keys): void
    {
        $this->keys = $keys;
    }


    /**
     * @return int
     */
    public function getStartRow(): int
    {
        return $this->startRow;
    }

    /**
     * @param  int  $startRow
     */
    public function setStartRow(int $startRow): void
    {
        $this->startRow = $startRow;
    }

    /**
     * @return int
     */
    public function getColumnCount(): int
    {
        return $this->columnCount;
    }

    /**
     * @param  int  $columnCount
     */
    public function setColumnCount(int $columnCount): void
    {
        $this->columnCount = $columnCount;
    }
}