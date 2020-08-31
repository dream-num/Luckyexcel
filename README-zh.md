简体中文 | [English](./README.md)

## 介绍
Luckyexcel，是一个适配 [Luckysheet](https://github.com/mengshukeji/Luckysheet) 的excel导入导出库，只支持.xlsx格式文件（不支持.xls）。

## 演示
[Demo](https://mengshukeji.github.io/LuckysheetDemo/)

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

## 环境
[Node.js](https://nodejs.org/en/) Version >= 6 

## 安装
```
npm install -g gulp-cli
npm install
gulp
```

## 用法（改进中）

#### 第一步
`gulp build`后`dist`文件夹下的bundle.js复制到项目目录，bundle.js即为项目核心代码

#### 第二步

导入bundle.js,界面上指定一个文件上传组件，编写类似如下的监听方法，调用`LuckyExcel.transformExcelToLucky`，然后在回调中获取到转换后的JSON数据，此JSON数据即是Luckysheet可识别的格式，使用Luckysheet初始化即可。
```js
function demoHandler(){
    let upload = document.getElementById("Luckyexcel-demo-file");
    if(upload){
        
        window.onload = () => {
            
            upload.addEventListener("change", function(evt){
                var files:FileList = (evt.target as any).files;
                LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any){

                    window.luckysheet.destroy();
                    
                    window.luckysheet.create({
                        container: 'luckysheet', //luckysheet is the container id
                        data:exportJson.sheets,
                        title:exportJson.info.name,
                        userInfo:exportJson.info.name.creator
                    });
                });
            });
        }
    }
}
```

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

## 版权信息
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, mengshukeji
