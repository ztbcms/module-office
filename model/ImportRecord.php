<?php
/**
 * Created by PhpStorm.
 * User: zhlhuang
 * Date: 2021/7/13
 * Time: 13:49.
 */

declare(strict_types=1);

namespace app\office\model;


use think\Model;

class ImportRecord extends Model
{
    protected $name = 'office_import_record';
    protected $pk = 'import_record_id';
}