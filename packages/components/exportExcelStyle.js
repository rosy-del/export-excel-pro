// ref获取到el-table的dom
// headerStyle 表头样式
//cellStyle 单元格样式
// name 文件名字
import ExcelJS from 'exceljs';
export async function exportExcelStyle(tableDom, headerStyle, cellStyle, name) {
    //    console.log(111);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    let eachHeader = [];
    let firstHeader = [];//第一次获得的表头
    // 拿到表头
    // 处理每一层表头 凡是colspan = 1就存入，不然就循环coslspan次存入
    //第二层表头 凡是colspan = 1就存入，不然就循环coslspan次存入
    const headerRows = tableDom.querySelectorAll('.el-table__header-wrapper > table > thead > tr');
    headerRows.forEach((headerRow, index) => {
        const cells = headerRow.querySelectorAll('th');
        eachHeader = [];
        cells.forEach((cell) => {
            // 如果colspan为1，则直接push，如果不为1，则循环colspan次，依次push
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            const cellText = cell.textContent;
            if (colspan === 1) {
                eachHeader.push(cellText);
            } else {
                for (let i = 0; i < colspan; i++) {
                    eachHeader.push(cellText);
                }
            }

            // console.log(eachHeader);
        });
        firstHeader.push(eachHeader)
    });
    // console.log(firstHeader);

    // 删除数组最后一个空白值
    firstHeader.forEach((item, index) => {
        item.pop()
    });
    // console.log(firstHeader);
    // 处理表头，遍历到每一层表头，就对比上一层表头的值，如果没有重复的值就将对应的
    // 置为‘’，如果有重复的就从第二层表头中取出值，并删除该数组这个值
    //第一层表头不需要处理，因为没有上一层表头
    // const headers = [];
    // const centerHeader = [];
    // const headerFirst = firstHeader[0];
    // for (let i = 0; i < firstHeader.length - 1; i++) {
    //     for (let j = 0; j < headerFirst.length; j++) {
    //         if (j < headerFirst.length && firstHeader[i][j] === firstHeader[i][j - 1]) {
    //             let endHeader = firstHeader[i + 1][j]
    //             let valueToMove = endHeader.shift();
    //             centerHeader[j - 1] = valueToMove;
    //         } else if (j < headerFirst.length && firstHeader[i][j] === firstHeader[i][j - 1] && firstHeader[i][j - 1] === firstHeader[i][j - 2]) {
    //             let endHeader = firstHeader[i + 1][j]
    //             let valueToMove = endHeader.shift();
    //             centerHeader[j - 1] = valueToMove;
    //             // 不重复的
    //         } else {
    //             centerHeader[j - 1] = '';
    //         }
    //     }
    //     headers.push(centerHeader)
    // }
    // console.log(headers);
    // const headers = [];
    const centerHeader = [];
    const headerFirst = firstHeader[0];
    for (let i = 0; i < firstHeader.length - 1; i++) {
        let endHeader = firstHeader[i + 1];
        // console.log(endHeader);       
        for (let j = 0; j < headerFirst.length; j++) {

            if (j < headerFirst.length && firstHeader[i][j] === firstHeader[i][j - 1] && emptyCheck(firstHeader[i][j]) && emptyCheck(firstHeader[i][j - 1])) {
                // console.log('相等',firstHeader[i][j],firstHeader[i][j-1],endHeader);         
                let valueToMove = endHeader.shift();
                centerHeader[j - 1] = valueToMove;
            } else if (j < headerFirst.length && firstHeader[i][j] !== firstHeader[i][j - 1] && firstHeader[i][j - 1] === firstHeader[i][j - 2]) {
                // 检查endHeader是否为数组且有元素可移除
                if (emptyCheck(firstHeader[i][j]) && emptyCheck(firstHeader[i][j - 1]) && emptyCheck(firstHeader[i][j - 1]) && emptyCheck(firstHeader[i][j - 2])) {
                    let valueToMove = endHeader.shift();
                    centerHeader[j - 1] = valueToMove;
                } else {
                    centerHeader[j - 1] = '';
                }
                // 不重复的置为空
            } else {
                centerHeader[j - 1] = '';
            }
        }
        // console.log(Object.values(centerHeader));

        firstHeader[i + 1] = Object.values(centerHeader);

    }

    //不包括第一层表头
    // console.log(firstHeader);

    // 将表头遍历到worksheet中
    for (let i = 1; i <= firstHeader.length; i++) {
        const row = worksheet.addRow(firstHeader[i - 1]);
        // 遍历当前行的每个单元格并应用表头样式
        row.eachCell((cell) => {
            cell.border = headerStyle.border;
            cell.alignment = headerStyle.alignment;
        });
    }
    //依靠数组长度生成excel列名
    let colNames = generateExcelColNames(firstHeader[0].length);
    // console.log(colNames);
    // 横向合并单元格(除了最后一层)
    for (let i = 0; i < firstHeader.length; i++) {
        let arr = [null, null];
        for (let j = 1; j <= firstHeader[i].length; j++) {
            if (firstHeader[i][j - 1] === firstHeader[i][j]) {    // 前后元素相同，表示可以合并
                if (!emptyCheck(arr[0]) && emptyCheck(firstHeader[i][j - 1])) {
                    arr[0] = colNames[j - 1] + (i + 1);
                }
                arr[1] = colNames[j] + (i + 1);
            } else {    // 前后元素不相同，j是从1开始，所以表示没有相同列或者此次相同列已结束
                if (emptyCheck(arr[0]) && emptyCheck(arr[1])) {
                    // arr[0]或者arr[1]为空，表示相邻元素不同，均有值表示有相同列，arr[1]便是最后一个相同的列
                    worksheet.mergeCells(arr[0] + ":" + arr[1]);
                }
                arr = [null, null];    // 相邻元素不同，arr重置，准备下一批可合并元素
            }
        }
    }

    // 纵向合并单元格
    // 第一层循环具体元素
    for (let i = 0; i < firstHeader[0].length; i++) {
        let sd = ""; // 开始元素
        let ed = ""; // 结束元素
        // 第二层循环，比较层级不同下标相同的元素
        for (let j = 1; j < firstHeader.length; j++) {
            if (firstHeader[j][i] === "") {  // 元素为空，表示可与上层元素合并
                sd = emptyCheck(sd) ? sd : colNames[i] + j;
                ed = colNames[i] + (j + 1);
            }
        }
        //合并并且再次垂直居中
        if (emptyCheck(sd) && emptyCheck(ed)) {
            worksheet.mergeCells(sd + ":" + ed);
        }
    }
    // 取出表格数据
    const elTablerow = [];
    const dataRows = tableDom.querySelectorAll('.el-table__body-wrapper >.el-table__body > tbody >.el-table__row');
    dataRows.forEach((dataRow) => {
        const rowData = [];
        const cells = dataRow.querySelectorAll('td > div');
        cells.forEach((cell) => {
            const cellText = cell.textContent;
            rowData.push(cellText);
        });

        elTablerow.push(rowData);
        // console.log(elTablerow);
    });
    // 获取表格数据
    const tableData = elTablerow;
    tableData.forEach((rowData) => {
        // console.log(rowData);
        const row = worksheet.addRow(rowData);
        row.eachCell((cell) => {
            cell.border = cellStyle.border;
            cell.alignment = cellStyle.alignment;
        });
    });
    // 设置列宽（根据需要调整列宽值）
    worksheet.columns.forEach((column) => {
        // console.log(column);
        column.width = 15;
        column.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    // 生成Excel文件并提供下载
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    // console.log(name);

    link.download = `${name}`;
    link.click();
    window.URL.revokeObjectURL(url);

}


// 公共函数

//获取数据类型
const getDataType = function (val) {
    return Object.prototype.toString
        .call(val)
        .replace(/^\[object (\S+)\]$/, "$1")
        .toLowerCase();
};
//非空判断

const emptyCheck = function (val) {
    if (val === null || val === "" || val === undefined) {
        return false;
    }
    let dataType = getDataType(val);
    switch (dataType) {
        case "string":
            return val.trim().length !== 0;
        case "array":
            return val.length !== 0;
        case "object":
            return Object.keys(val).length !== 0;
    }
    return true;
};




// 生成excel列名
function generateExcelColNames(length) {
    const colNames = [];
    let currentIndex = 0;
    if (length < 27) {
        // 生成单个字母的列名（A - Z）
        // 先处理单个字母的列名（A - Z）
        for (let i = 65; i <= 90; i++) {
            colNames.push(String.fromCharCode(i));
            currentIndex++;
            if (currentIndex === length) return colNames;
        }
    } else if (26 < length < 703) {
        // 先处理单个字母的列名（A - Z）
        for (let i = 65; i <= 90; i++) {
            colNames.push(String.fromCharCode(i));
            currentIndex++;
            if (currentIndex === length) return colNames;
        }
        // 再处理两个字母组合的列名（AA - ZZ）
        for (let i = 0; i < 26; i++) {
            for (let j = 0; j < 26; j++) {
                colNames.push(String.fromCharCode(65 + i) + String.fromCharCode(65 + j));
                currentIndex++;
                if (currentIndex === length) return colNames;
            }
        }
    } else if (length > 702) {
        // 先处理单个字母的列名（A - Z）
        for (let i = 65; i <= 90; i++) {
            colNames.push(String.fromCharCode(i));
            currentIndex++;
            if (currentIndex === length) return colNames;
        }
        // 再处理两个字母组合的列名（AA - ZZ）
        for (let i = 0; i < 26; i++) {
            for (let j = 0; j < 26; j++) {
                colNames.push(String.fromCharCode(65 + i) + String.fromCharCode(65 + j));
                currentIndex++;
                if (currentIndex === length) return colNames;
            }
        }
        for (let i = 0; i < 26; i++) {
            for (let j = 0; j < 26; j++) {
                for (let k = 0; k < 26; k++) {
                    colNames.push(String.fromCharCode(65 + i) + String.fromCharCode(65 + j) + String.fromCharCode(65 + k));
                    currentIndex++;
                    if (currentIndex === length) return colNames;
                }
            }
        }
        return colNames;
    }
}