SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for cms_office_import_record
-- ----------------------------
DROP TABLE IF EXISTS `cms_office_import_record`;
CREATE TABLE `cms_office_import_record`
(
    `import_record_id` int(11) unsigned NOT NULL AUTO_INCREMENT,
    `file_path`        varchar(128) DEFAULT '' COMMENT '导入文件地址',
    `data_md5`         varchar(128) DEFAULT NULL COMMENT '导入数据标识',
    `status`           tinyint(1) DEFAULT NULL COMMENT '导入状态0失败，1成功',
    `create_time`      int(11) DEFAULT '0' COMMENT '创建时间',
    `update_time`      int(11) DEFAULT '0' COMMENT '更新时间',
    PRIMARY KEY (`import_record_id`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ROW_FORMAT=COMPACT;
SET FOREIGN_KEY_CHECKS = 1;
