<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/2/18
 * Time: 09:45.
 */

namespace app\office\service;


use app\common\service\BaseService;
use app\office\service\filter\ExcelRowFilter;
use app\office\service\filter\ExcelValueFilter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use think\facade\Filesystem;

class ExcelService extends BaseService
{
    protected $fileName = '';
    protected $header = [];
    protected $data = [];
    protected $rowHeight = 20;
    protected $autoSize = false;

    /**
     * ExcelService constructor.
     * @param  string  $fileName
     * @param  array  $header
     * @param  array  $data
     */
    public function __construct(string $fileName = '', array $header = [], array $data = [])
    {
        $this->fileName = $fileName;
        $this->header = $header;
        $this->data = $data;
    }

    /**
     * 获取列数，获取A、B、C、AA
     * @param $i
     * @return string
     */
    private function getChar($i)
    {
        $y = ($i / 26);
        if ($y >= 1) {
            $y = intval($y);
            return chr($y + 64).chr($i - $y * 26 + 65);
        } else {
            return chr($i + 65);
        }
    }

    /**
     * 获取默认配置
     * @return array
     */
    private function getStyleArray()
    {
        return [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'borders'   => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE,
                ],
            ]
        ];
    }

    /**
     * 打印表头信息
     * @param $spreadsheet
     * @param $sheet
     * @return bool
     * @throws \Throwable
     */
    private function setPrintHeader($spreadsheet, $sheet)
    {
        $styleArray = $this->getStyleArray();
        $hasHeader = false;
        //打印表头
        if (!empty($this->header)) {
            $hasHeader = true;
            $i = 0;
            foreach ($this->header as $k => $v) {
                $spreadsheet->getActiveSheet()->getColumnDimension($this->getChar($i))->setAutoSize($this->autoSize);
                //设置默认高度
                $sheet->getRowDimension(1)->setRowHeight($this->rowHeight);
                $cellKey = $this->getChar($i).'1';
                throw_if(is_array($v), new \Exception('header value is not array'));

                $sheet->setCellValue($cellKey, $v);
                $sheet->getStyle($cellKey)->applyFromArray(array_merge([
                    'font' => [
                        'bold' => true,
                    ],
                ], $styleArray));
                $i++;
            }
        }
        return $hasHeader;
    }

    /**
     * @return Spreadsheet
     * @throws \Throwable
     */
    private function setCellValue()
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        //设置表头
        $hasHeader = $this->setPrintHeader($spreadsheet, $sheet);
        $styleArray = $this->getStyleArray();
        //数据为空
        throw_if(empty($this->data), new \Exception('data is not empty'));
        //有标头第一个从2开始
        $row = $hasHeader ? 2 : 1;
        foreach ($this->data as $key => $value) {
            $i = 0;
            //设置默认高度
            if ($this->rowHeight > 0) {
                $sheet->getRowDimension($row)->setRowHeight($this->rowHeight);
            } else {
                $sheet->getRowDimension($row)->setZeroHeight(true);
            }

            $mergeRow = 1;
            if ($value instanceof ExcelRowFilter) {
                $mergeRow = $value->getRow();
                $value = $value->getData();
            }
            foreach ($value as $k => $v) {
                $cellKey = $this->getChar($i).$row;
                $unmergeRow = 1;
                if ($v instanceof ExcelValueFilter) {
                    $unmergeRow = $v->getRow();
                    $v = $v->getFilterValue();
                    //这是记录中多行显示的数据
                    if ($unmergeRow > 1 && is_array($v)) {
                        foreach ($v as $kk => $vv) {
                            $unmergeCellKey = $this->getChar($i).($row + $kk);
                            $sheet->setCellValue($unmergeCellKey, $vv);
                        }
                    } elseif (is_array($v)) {
                        $sheet->setCellValue($cellKey, json_encode($v));
                    } else {
                        $sheet->setCellValue($cellKey, $v);
                    }
                } else {
                    if (is_array($v)) {
                        $sheet->setCellValue($cellKey, json_encode($v));
                    } else {
                        $sheet->setCellValue($cellKey, $v);
                    }
                }
                // 如果记录整体占用行数>1,且非分行显示数据
                if ($mergeRow > 1 && $unmergeRow == 1) {
                    //如果行数大于1，则合并行数
                    $mergeCellKey = $this->getChar($i).($row + ($mergeRow - 1));
                    $sheet->mergeCells($cellKey.':'.$mergeCellKey);
                }
                $sheet->getStyle($cellKey)->applyFromArray($styleArray);
                $i++;
            }
            $row += $mergeRow;
        }
        return $spreadsheet;
    }

    /**
     * @param  string  $fileType
     * @return string[]
     * @throws \Throwable
     */
    public function save($fileType = 'xls')
    {
        if (!$this->fileName) {
            //如果没有传文件名称，默认
            $this->fileName = date("Y-m-d H:i");
        }
        $saveTempPath = Filesystem::disk('public')->path('temp').'/';
        if (!is_dir($saveTempPath)) {
            mkdir($saveTempPath, 755);
        }

        $fileUrl = Filesystem::disk('public')->getConfig()->get('url').'/temp/';
        $fileName = $this->fileName.'.xls';
        $spreadsheet = $this->setCellValue();

        $writer = new Xls($spreadsheet);
        if (file_exists($saveTempPath.$fileName)) {
            // 已有文件删除覆盖
            unlink($saveTempPath.$fileName);
        }
        $writer->save($saveTempPath.$fileName);
        return [
            'url' => $fileUrl.$fileName
        ];
    }


    /**
     * @return string
     */
    public function getFileName(): string
    {
        return $this->fileName;
    }

    /**
     * @param  string  $fileName
     */
    public function setFileName(string $fileName): void
    {
        $this->fileName = $fileName;
    }


    /**
     * @return array
     */
    public function getHeader(): array
    {
        return $this->header;
    }

    /**
     * @param  array  $header
     */
    public function setHeader(array $header): void
    {
        $this->header = $header;
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
     * @return int
     */
    public function getRowHeight(): int
    {
        return $this->rowHeight;
    }

    /**
     * @param  int  $rowHeight
     */
    public function setRowHeight(int $rowHeight): void
    {
        $this->rowHeight = $rowHeight;
    }

    /**
     * @return bool
     */
    public function isAutoSize(): bool
    {
        return $this->autoSize;
    }

    /**
     * @param  bool  $autoSize
     */
    public function setAutoSize(bool $autoSize): void
    {
        $this->autoSize = $autoSize;
    }

}