# export-excel-pro（vue2+excelJs+el-table）多层表头导出

#### 直接从el-tabe表格渲染好的内容中取数！！！（包括表头和表体）

#### 封装好的excelJs导出

#### 所见即所得！！

#### 1.安装elementUI/elementplus

```js
npm install element-plus --save
npm install exceljs
```

#### 2.在main.js导入

```js
import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import App from './App.vue'

const app = createApp(App)

app.use(ElementPlus)
app.mount('#app')
```

#### 3.在哪用就在哪导入

```js
import ExcelJS from 'exceljs';
```

#### 4.设置表头和表格单元格样式

```js
//alignment、border都是excelJs官方的样式属性，可见中文文档
//https://gitee.com/alan_scut/exceljs#styles
//numFmt、font 、alignment 、border、 fill
例子：
// 设置表头样式
      const headerStyle = {
        alignment: {
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: 'black' },
          bottom: { style: 'thin', color: 'black' },
          left: { style: 'thin', color: 'black' },
          right: { style: 'thin', color: 'black' }
        }
      }
      // 设置普通单元格样式
      const cellStyle = {
        alignment: {
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: 'black' },
          bottom: { style: 'thin', color: 'black' },
          left: { style: 'thin', color: 'black' },
          right: { style: 'thin', color: 'black' }
        }
      }
```

#### 5.*ref获取到el-table的dom*

```js
const tableDom = this.$refs['report-table'].$el;
//el-table上面的ref名字为report-table
```

#### 6.文件名字

```js
 const name = 'exported_file_example';
```

#### 7.例子

```js
 exportExcel() {
      // 设置表头样式
      const headerStyle = {
        alignment: {
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: 'black' },
          bottom: { style: 'thin', color: 'black' },
          left: { style: 'thin', color: 'black' },
          right: { style: 'thin', color: 'black' }
        }
      }
      // 设置普通单元格样式
      const cellStyle = {
        alignment: {
          horizontal: 'center',
          vertical: 'center'
        },
        border: {
          top: { style: 'thin', color: 'black' },
          bottom: { style: 'thin', color: 'black' },
          left: { style: 'thin', color: 'black' },
          right: { style: 'thin', color: 'black' }
        }
      }
      const tableDom = this.$refs['report-table'].$el;
      exportExcelStyle(tableDom, headerStyle, cellStyle, name);
    }
```

