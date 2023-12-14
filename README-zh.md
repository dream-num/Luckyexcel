简体中文 | [English](./README.md)

## 介绍
Luckyexcel，是一个适配 [Luckysheet](https://github.com/mengshukeji/Luckysheet) 的excel导入导出库，只支持.xlsx格式文件（不支持.xls）。

## 演示
[Demo](https://mengshukeji.github.io/LuckyexcelDemo/)

## 特性
支持excel文件导入到Luckysheet适配列表

- 单元格样式
- 单元格边框
- 单元格格式，如数字格式、日期、百分比等
- 公式

### 计划

目标是支持所有Luckysheet支持的特性

- 条件格式
- 数据透视表
- 图表
- 排序
- 筛选
- 批注
- excel导出

## 用法

### CDN
```html
<script src="https://cdn.jsdelivr.net/npm/luckyexcel/dist/luckyexcel.umd.js"></script>
<script>
    // 先确保获取到了xlsx文件file，再使用全局方法window.LuckyExcel转化
    LuckyExcel.transformExcelToLucky(
        file, 
        function(exportJson, luckysheetfile){
            // 获得转化后的表格数据后，使用luckysheet初始化，或者更新已有的luckysheet工作簿
            // 注：luckysheet需要引入依赖包和初始化表格容器才可以使用
            luckysheet.create({
                container: 'luckysheet', // luckysheet is the container id
                data:exportJson.sheets,
                title:exportJson.info.name,
                userInfo:exportJson.info.name.creator
            });
        },
        function(err){
            logger.error('Import failed. Is your fail a valid xlsx?');
        }
    );
</script>
```
> 案例 [Demo index.html](./src/index.html)展示了详细的用法

### ES 和 Node.js

#### 安装
```shell
npm install luckyexcel
```

#### ES导入
```js
import LuckyExcel from 'luckyexcel'

// 得到xlsx文件后
LuckyExcel.transformExcelToLucky(
    file, 
    function(exportJson, luckysheetfile){
        // 转换后获取工作表数据
    },
    function(error){
        // 如果抛出任何错误，则处理错误
    }
)
```
> 案例 [luckysheet-vue](https://github.com/mengshukeji/luckysheet-vue)

#### Node.js导入
```js
var fs = require("fs");
var LuckyExcel = require('luckyexcel');

// 读取一个xlsx文件
fs.readFile("House cleaning checklist.xlsx", function(err, data) {
    if (err) throw err;

    LuckyExcel.transformExcelToLucky(data, function(exportJson, luckysheetfile){
        // 转换后获取工作表数据
    });

});
```
> 案例 [Luckyexcel-node](https://github.com/mengshukeji/Luckyexcel-node)


## 开发

### 环境
[Node.js](https://nodejs.org/en/) Version >= 6 

### 安装
```
npm install -g gulp-cli
npm install
```
### 开发
```
npm run dev
```
### 打包
```
npm run build
```

项目中使用了第三方插件：[JSZip](https://github.com/Stuk/jszip)，感谢！

## 交流
- 任何疑问或者建议，欢迎提交[Issues](https://github.com/mengshukeji/Luckyexcel/issues/)

- 添加小编微信,拉你进Luckysheet开发者交流微信群,备注:加群

  <img src="/docs/.vuepress/public/img/%E5%BE%AE%E4%BF%A1%E4%BA%8C%E7%BB%B4%E7%A0%81.jpg" width = "200" alt="微信群" align="center" />

- 加入Luckysheet开发者交流QQ群
  
  <img src="/docs/.vuepress/public/img/QQ%E7%BE%A4%E4%BA%8C%E7%BB%B4%E7%A0%81.jpg" width = "200" alt="微信群" align="center" />


## 贡献者和感谢
- [@wbfsa](https://github.com/wbfsa)
- [@wpxp123456](https://github.com/wpxp123456)
- [@Dushusir](https://github.com/Dushusir)
- [@xxxDeveloper](https://github.com/xxxDeveloper)

## 版权信息
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, mengshukeji
