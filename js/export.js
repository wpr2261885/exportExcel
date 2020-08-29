﻿// 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
function sheet2blob(sheet, sheetName) {
    // sheetName = sheetName || 'sheet1';
    // let workbook = {
    //     SheetNames: [sheetName],
    //     Sheets: {}
    // };
    // workbook.Sheets[sheetName] = sheet; // 生成excel的配置项

    let workbook = {
        SheetNames: [],
        Sheets: {}
    };
    for (let i = 0; i < sheet.length; i++) {
        if (workbook.SheetNames.indexOf(sheet[i].name) > -1) {
            sheet[i].name += '(1)'
        }
        workbook.SheetNames.push(sheet[i].name)
        workbook.Sheets[sheet[i].name] = sheet[i].sheet; // 生成excel的配置项
    }

    let wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    let wbout = XLSX.write(workbook, wopts);
    let blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    }); // 字符串转ArrayBuffer
    function s2ab(s) {
        let buf = new ArrayBuffer(s.length);
        let view = new Uint8Array(buf);
        for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}

function openDownloadDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    let aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    let event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}
