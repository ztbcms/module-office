<?php
/**
 * Author: cycle_3
 */

namespace app\office\service;

use app\common\service\BaseService;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\PhpWord;

class WordService extends BaseService
{

    protected $fileName = '';
    protected $data = [];

    /**
     * ExcelService constructor.
     * @param  string  $fileName
     * @param  array  $header
     * @param  array  $data
     */
    public function __construct($fileName = '',  array $data = [])
    {
        $this->fileName = $fileName;
        $this->data = $data;
    }

    /**
     * 替换模板内容
     * @param string $template_url
     */
    public function replaceWord($template_url = ''){
        //指定事先制作好的模板文件路径
        $template_document = new TemplateProcessor($template_url);
        foreach ($this->data as $k => $v) {
            $template_document->setValue($k, $v);
        }
        $template_document->saveAs( public_path("word/demo").$this->fileName.'.docx');
        $file_url = public_path("word/demo").$this->fileName.'.docx';
        $this->downloadFile($file_url);
    }

    /**
     * 创建work文件
     */
    public function createWord(){
        $phpWord = new PhpWord();

        //默认字体
        if(isset($this->data['defaultFontName'])) {
            $phpWord->setDefaultFontName($this->data['defaultFontName']);
        } else {
            $phpWord->setDefaultFontName('Tahoma');
        }

        //默认大小
        if(isset($this->data['defaultFontSize'])) {
            $phpWord->setdefaultFontSize($this->data['defaultFontSize']);
        } else {
            $phpWord->setDefaultFontSize(12);
        }

        $section = $phpWord->addSection();
        foreach ($this->data['content'] as $k => $v) {
            if($v['type'] == 'text') {
                $section->addText( $v['val'], $v['font_style'] ,$v['paragraph_style']);
            } else if($v['type'] == 'img') {
                $section->addImage($v['val'], $v['img_style']);
            } else if($v['type'] == 'form') {
                $table = $section->addTable($v['style']);
                foreach ($v['val'] as $k2 => $v2) {
                    $table->addRow($v2['height']);
                    foreach ($v2['content'] as $k3 => $v3) {
                        $table->addCell($v3['width'],$v3['cell_style'])->addText($v3['val']);
                    }
                }
            } else if($v['type'] == 'line') {
                if(empty($v['val'])) {
                    $section->addTextBreak();
                } else {
                    $section->addTextBreak($v['val']);
                }
            }
        }

        //生成Word文档
        $filePath = public_path("word/demo").$this->fileName.'.docx';
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($filePath);

        $this->downloadFile($filePath);
    }

    /**
     * 下载文档操作
     * @param $file_url
     * @param null $new_name
     */
    function downloadFile($file_url, $new_name = null)
    {
        if (!isset($file_url) || trim($file_url) == '') {
            echo '500';
        }
        if (!file_exists($file_url)) { //检查文件是否存在
            echo '404';
        }
        $file_name = basename($file_url);
        $file_type = explode('.', $file_url);
        $file_type = $file_type[count($file_type) - 1];
        $file_name = trim($new_name == '') ? $file_name : urlencode($new_name);
        $file_type = fopen($file_url, 'r'); //打开文件
        //输入文件标签
        header("Content-type: application/octet-stream");
        header("Accept-Ranges: bytes");
        header("Accept-Length: " . filesize($file_url));
        header("Content-Disposition: attachment; filename=" . $file_name);
        //输出文件内容
        echo fread($file_type, filesize($file_url));
        fclose($file_type);
    }

}