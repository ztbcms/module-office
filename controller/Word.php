<?php
/**
 * Author: cycle_3
 */

namespace app\office\controller;

use app\common\controller\AdminController;
use app\office\service\WordService;

/**
 * 导出文档
 * Class Word
 * @package app\office\controller
 */
class Word extends AdminController
{

    /**
     * 替换模板案例
     */
    public function replaceWord()
    {
        $content = [
            'first_name'           => '马',
            'last_name'            => '云',
            'birth'                => '2012/05/22',
            'phone'                => '10086',
            'emergency_name'       => '马化腾',
            'emergency_phone'      => '10010',
            'allergy'              => '微信',
            'family_doctor_name'   => '雷军',
            'family_doctor_phone'  => '13800138000',
            'family_doctor_clinic' => '小米手机',
            'health_precautions'   => '客户经理',
            'creator_name'         => '吴一平',
            'creator_phone'        => '1588744',
        ];
        $WordService = new WordService('导出替换模板', $content);
        $directory = app_path().'demo'.DIRECTORY_SEPARATOR.'word'.DIRECTORY_SEPARATOR;
        $template_url = $directory.'demo.docx';
        try {
            $file_url = $WordService->replaceWord($template_url)->getFileUrl();
            return self::makeJsonReturn(true, ['file_url' => $file_url]);
        } catch (\Exception $exception) {
            return self::makeJsonReturn(false, [], $exception->getMessage());
        }
    }


    /**
     * 创建work文档
     */
    public function createWord()
    {
        //具体调整可参考 https://segmentfault.com/a/1190000019479817?utm_source=tag-newest
        $content = [
            'defaultFontName' => 'Tahoma',  //默认使用的字体
            'defaultFontSize' => '12',   //默认使用的字号
            'content'         => [
                [
                    'type'            => 'text',  //文本类型
                    'val'             => '这是一个文本', //内容
                    'font_style'      => [
                        'bold'  => true,
                        'color' => 'AACC00',
                        'size'  => 18,
                        'align' => 'center',
                    ],
                    'paragraph_style' => [
                        'align'       => 5,   //水平对齐:leftrightcenterboth / justify
                        'spaceBefore' => 10   //段前间距，单位： twips.
                    ]
                ],
                [
                    'type' => 'line',  //换行类型
                    'val'  => 2
                ],
                [
                    'type'      => 'img',  //图片类型
                    'val'       => 'https://gzztb.coding.net/static/fruit_avatar/Fruit-1.png', //路径
                    'img_style' => [
                        'width'  => 350,
                        'height' => 350,
                        'align'  => 'center',
                    ],
                ],
                [
                    'type' => 'line',  //换行类型
                    'val'  => 2
                ],
                [
                    'type'  => 'form',  //表格类型
                    'val'   => [
                        [
                            'content' => [
                                [
                                    'val'        => 'Cell 1',
                                    'width'      => 1000,
                                    'cell_style' => [
                                        'valign' => 'center'
                                    ]
                                ],
                                [
                                    'val'        => 'Cell 2',
                                    'width'      => 3000,
                                    'cell_style' => [
                                        'valign' => 'center'
                                    ]
                                ],
                                [
                                    'val'        => 'Cell 3',
                                    'width'      => 2000,
                                    'cell_style' => [
                                        'valign' => 'center'
                                    ]
                                ]
                            ],
                            'height'  => 400
                        ],
                        [
                            'content' => [
                                [
                                    'val'        => 'Cell 4',
                                    'width'      => 1000,
                                    'cell_style' => [
                                        'valign' => 'center'
                                    ]
                                ],
                                [
                                    'val'        => 'Cell 5',
                                    'width'      => 3000,
                                    'cell_style' => [
                                        'valign' => 'center'
                                    ]
                                ],
                                [
                                    'val'        => 'Cell 6',
                                    'width'      => 2000,
                                    'cell_style' => [
                                        'valign' => 'center'
                                    ]
                                ]
                            ],
                            'height'  => 400
                        ],
                    ],
                    'style' => [
                        'borderColor' => '006699',
                        'borderSize'  => 6,
                        'cellMargin'  => 50,
                        'align'       => 'center'
                    ]
                ],
            ],
        ];
        $WordService = new WordService('新建Word模板', $content);
        try {
            $file_url = $WordService->createWord()->getFileUrl();
            return self::makeJsonReturn(true, ['file_url' => $file_url]);
        } catch (\Exception $exception) {
            return self::makeJsonReturn(false, [], $exception->getMessage());
        }
    }
}