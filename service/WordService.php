<?php
/**
 * Author: cycle_3
 */

namespace app\office\service;

use app\common\service\BaseService;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\PhpWord;
use think\facade\Filesystem;

class WordService extends BaseService
{

    protected $fileName = '';
    protected $data = [];
    protected $file_path = '';
    protected $file_url = '';

    /**
     *
     * WordService constructor.
     * @param  string  $fileName
     * @param  array  $data
     */
    public function __construct($fileName = '', array $data = [])
    {
        $this->fileName = $fileName;
        $this->data = $data;
    }

    /**
     * 替换模板内容
     * @param  string  $template_url
     * @return $this
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     */
    public function replaceWord($template_url = ''): WordService
    {
        //指定事先制作好的模板文件路径
        $template_document = new TemplateProcessor($template_url);
        foreach ($this->data as $k => $v) {
            $template_document->setValue($k, $v);
        }
        $this->file_path = Filesystem::disk('public')->path('temp').'/'.$this->fileName.'.docx';
        $this->file_url = Filesystem::disk('public')->getConfig()->get('url').'/temp/'.$this->fileName.'.docx';
        $template_document->saveAs($this->file_path);
        return $this;
    }

    /**
     * 下载文档操作
     * @param  null  $new_name
     */
    function downloadFile($new_name = null)
    {
        if (!isset($this->file_path) || trim($this->file_path) == '') {
            echo '500';
        }
        if (!file_exists($this->file_path)) { //检查文件是否存在
            echo '404';
        }
        $file_name = basename($this->file_path);
        $file_name = trim($new_name == '') ? $file_name : urlencode($new_name);
        $file_type = fopen($this->file_path, 'r'); //打开文件
        //输入文件标签
        header("Content-type: application/octet-stream");
        header("Accept-Ranges: bytes");
        header("Accept-Length: ".filesize($this->file_path));
        header("Content-Disposition: attachment; filename=".$file_name);
        //输出文件内容
        echo fread($file_type, filesize($this->file_path));
        fclose($file_type);
    }


    /**
     * 创建work文件
     * @return $this
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function createWord(): WordService
    {
        $phpWord = new PhpWord();

        //默认字体
        if (isset($this->data['defaultFontName'])) {
            $phpWord->setDefaultFontName($this->data['defaultFontName']);
        } else {
            $phpWord->setDefaultFontName('Tahoma');
        }

        //默认大小
        if (isset($this->data['defaultFontSize'])) {
            $phpWord->setdefaultFontSize($this->data['defaultFontSize']);
        } else {
            $phpWord->setDefaultFontSize(12);
        }

        $section = $phpWord->addSection();
        foreach ($this->data['content'] as $k => $v) {
            if ($v['type'] == 'text') {
                $section->addText($v['val'], $v['font_style'], $v['paragraph_style']);
            } else {
                if ($v['type'] == 'img') {
                    $section->addImage($v['val'], $v['img_style']);
                } else {
                    if ($v['type'] == 'form') {
                        $table = $section->addTable($v['style']);
                        foreach ($v['val'] as $k2 => $v2) {
                            $table->addRow($v2['height']);
                            foreach ($v2['content'] as $k3 => $v3) {
                                $table->addCell($v3['width'], $v3['cell_style'])->addText($v3['val']);
                            }
                        }
                    } else {
                        if ($v['type'] == 'line') {
                            if (empty($v['val'])) {
                                $section->addTextBreak();
                            } else {
                                $section->addTextBreak($v['val']);
                            }
                        }
                    }
                }
            }
        }

        //生成Word文档
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $this->file_path = Filesystem::disk('public')->path('temp').'/'.$this->fileName.'.docx';
        $this->file_url = Filesystem::disk('public')->getConfig()->get('url').'/temp/'.$this->fileName.'.docx';
        $objWriter->save($this->file_path);
        return $this;
    }

    /**
     * @return string
     */
    public function getFileUrl(): string
    {
        return $this->file_url;
    }
}