<div>
    <div id="app" style="padding: 8px;" v-cloak>
        <div>
            <el-card>
                <h3>导出示例</h3>
                <div>
                    <el-button @click="exportEvent" type="primary">点击导出Xls</el-button>
                </div>
            </el-card>
        </div>
    </div>
    <script>
        $(document).ready(function () {
            new Vue({
                el: '#app',
                data: {},
                methods: {
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
                },
            })
        })
    </script>
</div>