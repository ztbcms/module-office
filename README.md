### 文档管理

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

#### 更新日志

##### 0218
1. Excel 导出：行过滤、值过滤