<div>
    <div id="app" style="padding: 8px;" v-cloak>
        <div>
            <el-card>
                <h3>导出示例</h3>
                <div>
                    <el-button @click="exportEvent" type="primary">点击导出Xls</el-button>
                </div>
                <div style="margin-top: 10px">
                    <el-button type="primary" @click="gotoUploadFile">上传文件</el-button>
                    <el-link type="primary"
                             href="http://ztbcms-tes.oss-cn-beijing.aliyuncs.com/file/20210219/52134b2a202f6271da4b22b6b818d1e6.xlsx"
                             target="_blank">下载模板
                    </el-link>
                    <span style="padding-left: 20px;color: red;">ps:建议使用本地上传驱动</span>
                </div>
                <div v-if="select_file.aid" style="margin-top: 10px">
                    {{select_file.filepath}}
                    <div style="margin-top: 10px">
                        <el-button type="primary" size="small" @click="importFile">确认导入</el-button>
                    </div>
                    <div>
                        <pre>{{import_result}}</pre>
                    </div>
                </div>
            </el-card>
        </div>
    </div>
    <script>
        $(document).ready(function () {
            new Vue({
                el: '#app',
                data: {
                    select_file: {},
                    import_result: ""
                },
                methods: {
                    importFile: function () {
                        if (!this.select_file.filepath) {
                            this.$message.error('请选择导入文件');
                            return
                        }
                        var _this = this
                        this.httpGet("{:api_url('office/index/importXls')}", {
                            filepath: this.select_file.filepath
                        }, function (res) {
                            if (res.status) {
                                _this.import_result = res.data
                            } else {
                                _this.$message.error(res.msg);
                            }
                        })
                    },
                    gotoUploadFile: function () {
                        layer.open({
                            type: 2,
                            title: '',
                            closeBtn: false,
                            content: "{:api_url('common/upload.panel/fileUpload')}",
                            area: ['720px', '550px'],
                        })
                    },
                    onUploadedFile: function (event) {
                        var files = event.detail.files;
                        if (files[0]) {
                            console.log('files[0]', files[0])
                            this.select_file = files[0]
                        }
                    },
                    exportEvent: function () {
                        this.httpGet("{:api_url('office/index/exportXls')}", {}, function (res) {
                            console.log(res)
                            if (res.status) {
                                var url = res.data.url
                                location.href = url
                            }
                        })
                    }
                },
                mounted: function () {
                    window.addEventListener('ZTBCMS_UPLOAD_FILE', this.onUploadedFile.bind(this));
                },
            })
        })
    </script>
</div>