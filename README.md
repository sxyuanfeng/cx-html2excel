[![npm version](https://badge.fury.io/js/cx-html2excel.svg)](https://www.npmjs.com/package/cx-html2excel)

# cx-html2excel
一个将html的table标签导出成excel的库，依赖exceljs，支持各种样式、对齐、图片、合并单元格、颜色、边框、单元格格式等

Usage
-----
1. npm
```shell
npm install cx-html2excel file-saver exceljs
```

```ts
import { generateExcel } from 'cx-html2excel';
import { saveAs } from 'file-saver';

async function handleExport() {
    let table_elt = document.getElementById('table-id');
    let buffer = await generateExcel([{name: 'sheetName', elet: table_elt}], {UTC: true});
    saveAs(
        new Blob([buffer], { type: 'application/vnd.openxmlformats' }),
        'excel.xlsx'
    );
}
```

2. script
```html
<script src="demo/exceljs.bare.js"></script>
<script src="demo/FileSaver.js"></script>
<script src="dist/cx-html2excel.js"></script>
<script>
    async function handleExport() {
        let table_elt = document.getElementById('table-id');
        let buffer = await html2excel.generateExcel([{name: 'sheetName', elet: table_elt}], { UTC: true });
        saveAs(
            new Blob([buffer], { type: 'application/vnd.openxmlformats' }),
            'excel.xlsx'
        );
    }
</script>
```

DOM
```html
<table id='table-id' style="border-collapse: collapse; border-spacing: 0;">
    <colgroup>
        <col style='width: 191px' />
        <col style='width: 299px' />
        <col style='width: 155px' />
        <col style='width: 200px' />
        <col style='width: 174px' />
        <col style='width: 172px' />
        <col style='width: 167px' />
        <col style='width: 200px' />
        <col style='width: 167px' />
        <col style='width: 200px' />
        <col style='width: 119px' />
        <col style='width: 119px' />
    </colgroup>
    <tr style='height: 26px'></tr>
    <tr>
        <td colSpan="2" style="border: 1px solid rgb(7, 83, 247); font-weight: 700;">
        螺丝起子
        </td>
        <td colSpan="3" style="border: 1px solid rgb(7, 83, 247); color: #abc; font-size: 8px">
        莫斯科骡子
        </td>
        <td colSpan="3" style="border: 1px solid rgb(7, 83, 247); color: rgb(237, 16, 42);">
        玛格丽特
        </td>
    </tr>
    <tr>
        <td style="border: 1px solid rgb(7, 83, 247)">
        伏特加
        </td>
        <td style="border: 1px solid rgb(247, 7, 95)">
        青柠汁
        </td>
        <td style="border: 1px solid rgb(247, 87, 7); text-align: center; color: aqua;">
        伏特加
        </td>
        <td style="border: 1px solid rgb(3, 8, 19); text-align: right; color: bisque;">
        橙汁
        </td>
        <td style="border: 1px solid rgb(247, 239, 7); text-align: left; color: chartreuse;">
        姜汁
        </td>
        <td style="border: 1px solid rgb(147, 7, 247)">
        龙舌兰
        </td>
        <td style="border: 1px solid rgb(7, 247, 23)">
        君度橙酒
        </td>
        <td style="border: 1px solid rgb(247, 7, 35)">
        青柠汁
        </td>
    </tr>
    <tr>
        <td colspan="2" data-z="0.00%" style="border-right: 1px solid #000">
        6.876%
        </td>
        <td colspan="1" data-z="_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)">
        31,243,256.067
        </td>
        <td colspan="1" data-v="whisky + vodka" style="border-top: 1px solid #000">
        oolong
        </td>
        <td colspan="2"  style="border-left: 1px solid #000">
        2024/03/14
        </td>
        <td data-t="s" style="border-bottom: 1px solid #000;">
        0984563
        </td>
        <td style="border-right: 1px solid #000;">
        0984563
        </td>
    </tr>
    <tr>
        <td colspan="8">
        <img src="./demo/logo192.png">
        </td>
    </tr>
</table>
```

API
---
```ts
generateExcel(
    tables: [{name: 'sheetName', elet: table_elt}],
    options?: {
        rows: true,
        font: true,
        alignment: true,
        border: true,
        numfmt: true,
        chart: true，
        dateNF: 'm/d/yy',
        UTC: false
    }): Promise<any>
```

