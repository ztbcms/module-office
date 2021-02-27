### 文档管理

#### 依赖

**Excel安装composer** `composer require phpoffice/phpspreadsheet`

**Word安装composer** `composer require phpoffice/phpword`

#### Excel 导出
- 初始化

```php
use app\office\service\ExcelService;
$ex = new ExcelService($fileName, $fileHeader, $exportData);
```

$fileName： 字符串 导出文件名称，为空则默认 按照时间格式 Y-m-d H:i

$fileHeader：数组 数据标题头显示  例如：["订单号", "订单总价", "订单商品", "下单时间"]

$exportData：数组 二维数据，可以是普通数组对象，也可以是相应的过滤器

```php 
     $exportData = [
            // 行过滤器
            new ExcelRowFilter(
                [
                    'a' => 1,
                    'b' => 1.1,
                    'c' => new GoodsFilter(['hello', 'world', 'excel'], 3), //值过滤器
                    'd' => new DateFilter(time())
                ], 3 
            ),
            [
                'a' => 2,
                'b' => 2.1,
                'c' => "hello 2",
                'd' => new DateFilter(time())
            ], [
                'a' => 3,
                'b' => 3.1,
                'c' => "hello 3",
                'd' => date('Y-m-d H:i:s', time())
            ]
        ];
```

- 行过滤器

定义：
```php 
new ExcelRowFilter(
                [
                    'a' => 1,
                    'b' => 1.1,
                    'c' => new GoodsFilter(['hello', 'world', 'excel'], 3),
                    'd' => new DateFilter(time())
                ], 3
            )
```
data：第一个参数一定是array格式

row：第二参数默认是1，这个代表整个记录占用行数，如果记录中含有类似商品这样需要多行显示的数据，则需要设置。
- 值过滤器

定义：
```
new GoodsFilter(['hello', 'world', 'excel'], 3)
```

value：第一个参数为展示数据的值，可以是字符串或数组

row：数据所占行数，如果>1 ,代表该数据需要多行显示，则value一定是array。

![img](https://s3.ax1x.com/2021/02/18/yWPnV1.png)

#### Excel 导入

- 导入模板

对于每一个导入行为都需要定一个导入模板，这个模板规定导入起始行`$startRow`、列数量`$columnCount`以及构建数组的关键字`$keys`

$startRow : 一般导入的xls文件都有一些表头信息，可通过起始行来忽略这些表头数据

$columnCount : 如果一个表有多余的列，可以通过限制列的数量来提供效率（因为可能在表制作、修改过程中会出现过多的空列的情况）

$keys : 每个列对应的key，可是字符串，也可以是"列过滤器"

```php
class DemoImportTemplate extends ExcelImportService
{
    protected $startRow = 4;
    protected $columnCount = 5;

    public function __construct(string $filePath)
    {
        parent::__construct($filePath);
        $this->keys = [
            'id', 'company_name', 'company_description', 'work_name',
            new ExcelColumnFilter('work_level', TrimFilter::class)
        ];
    }
}
```

- 列过滤
在我们导入数据过程中，会有一些数据格式问题（例如日期，空格、换行）这些都是肉眼看不到，到时会在导入后对数据计算处理造成影响。
在keys数组中，使用 `ExcelColumnFilter` 对象最为key 他构造函数有两个参数，

    - key 代表这个列关键字，可以字符串定义key相同

    - filter 过滤器，定义这个列使用的过滤器，如果定义了过滤器，则在数据读取后，会进入到过滤处理后返回

- 合并单元

问题：在示例给出的[文件](http://ztbcms-tes.oss-cn-beijing.aliyuncs.com/file/20210219/52134b2a202f6271da4b22b6b818d1e6.xlsx)中，一个单位可以招聘多个岗位，在第一、二列中，单位的名称和序号是合并的，如果正常获取会导致同一个单位除了第一个岗位外，其他的获取不到单位名称。

解决：我们在处理导入数据中，已经对合并数据进行了处理。具体看下面返回

```
[
    {
        "id": "1",
        "company_name": "江门市富华松苑管理处",
        "company_description": "江门市富华松苑管理处为江门市纪委监委属下正科级事业单位，经费按财政补助一类拨付，主要职能是承担富华松苑的日常管理、组织协调、后勤保障服务等工作。",
        "work_name": "职员",
        "work_level": "专业技术12级"
    },
    {
        "id": "2",
        "company_name": "江门市口腔医院",
        "company_description": "江门市口腔医院是具有六十多年历史的国家公立非营利医院，是市卫生健康局直属财政核补事业单位，拥有一批博士，硕士和高级职称口腔专业技术人才，是五邑地区口腔医疗，牙病防治，科研教育的中心和大中专教育的实习基地，被评为江门市文明单位和广东省文明窗口。下属九个门诊部，全部实行垂直统一管理，统一收费标准。口腔修复科是广东省科教兴医工程重点专科。",
        "work_name": "口腔医师",
        "work_level": "专业技术10级"
    },
    {
        "id": "2",
        "company_name": "江门市口腔医院",
        "company_description": "江门市口腔医院是具有六十多年历史的国家公立非营利医院，是市卫生健康局直属财政核补事业单位，拥有一批博士，硕士和高级职称口腔专业技术人才，是五邑地区口腔医疗，牙病防治，科研教育的中心和大中专教育的实习基地，被评为江门市文明单位和广东省文明窗口。下属九个门诊部，全部实行垂直统一管理，统一收费标准。口腔修复科是广东省科教兴医工程重点专科。",
        "work_name": "口腔医师",
        "work_level": "专业技术11级"
    }
]
```

#### 更新日志

##### 0219
1. Excel 导入：列过滤，合并单元

##### 0218
1. Excel 导出：行过滤、值过滤

##### 0227
1. Word 替换模板 
2. Word 创建自定义内容