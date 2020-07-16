import XlsxPopulate from "@/browser/xlsx-populate.min.js"
let optins = {
    mode: "edit", // edit | read
    showToolbar: true,
    showGrid: true,
    showContextmenu: true,
    view: {
        height() {
            return document.documentElement.clientHeight;
        },
        width() {
            return document.documentElement.clientWidth;
        }
    },
    rows: [],
    merges: [],
    styles: [
        {
            align: "center",
            border: {
                top: ["thin", "#000"],
                bottom: ["thin", "#000"],
                left: ["thin", "#000"],
                right: ["thin", "#000"]
            }
        }
    ],
};
let workbook
export default {
    async getWorkbook(filedata) {
        // 用浏览器的方式获取文件Blob
        return XlsxPopulate.fromDataAsync(filedata);
    },
    __computeMerges(merges) {
        let ASCII = [];
        for (let i = 65; i < 95; i++) {
            ASCII.push(String.fromCharCode(i));
        }
        let mergeskeys = Object.keys(merges);
        let keysList = mergeskeys.map(keys => keys.split(":"));
        return keysList.map(item => {
            if (item.length > 1) {
                let collect;
                item.forEach((value, index) => {
                    if (index == 0) {
                        // 处理开头
                        let R = value.charCodeAt(value.match(/\D+/g));
                        let C = Number.parseInt(value.match(/\d+/g));
                        collect = {
                            R,
                            C,
                            merge: [0, 0]
                        };
                    } else {
                        // 处理结尾    // ASCII
                        let R = value.charCodeAt(value.match(/\D+/g)) - collect.R;
                        let C = Number.parseInt(value.match(/\d+/g)) - collect.C;
                        collect.merge[1] = R;
                        collect.merge[0] = C;
                    }
                });
                collect.R = ASCII.indexOf(String.fromCharCode(collect.R));
                collect.C = collect.C - 1;
                return collect;
            }
            item = item.toString(); //处理单个
            return {
                R: ASCII.indexOf(item.match(/\D+/g).toString()),
                C: Number.parseInt(item.match(/\d+/g)) - 1,
                merge: [0, 0]
            };
        });
    },
    async loadExel(urlWorkbook, filedata) {
        urlWorkbook ? workbook = urlWorkbook : workbook = await this.getWorkbook(filedata);
        let sheet = workbook.sheet(0);
        let computedMerges = this.__computeMerges(sheet._mergeCells);
        let rows = sheet._rows;
        let merges = Object.keys(sheet._mergeCells);
        let row = {};
        row.len = rows.length + 10;
        rows.forEach((item, index) => {
            let cells = item._cells;
            row[index - 1] = {
                cells: {}
            };
            cells.forEach((cell, cellIndex) => {
                let cellMerge = computedMerges.find(item => {
                    if (item.C == index - 1 && item.R == cellIndex - 1) {
                        return item;
                    }
                });
                if (!cellMerge) {
                    cellMerge = { merge: [0, 0] };
                }
                row[index - 1].cells[cellIndex - 1] = {
                    text: cell._value,
                    merge: cellMerge.merge,
                    style: 0
                };
            });
        });
        optins.rows = row
        optins.merges = merges
        return optins
    },
    async generateBlob(newWorkbook) {
        let blob = await newWorkbook.outputAsync()
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveOrOpenBlob(blob, "out.xlsx");
        } else {
            let url = window.URL.createObjectURL(blob);
            let a = document.createElement("a");
            document.body.appendChild(a);
            a.href = url;
            a.download = "out.xlsx";
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        }
    },
    async exportFile(newExcel) {
        let { rows, merges, styles } = newExcel[0]
        let newWorkbook = await XlsxPopulate.fromBlankAsync()
        let sheet = newWorkbook.sheet(0)
        for (let columnIndex in rows) {
            if (!isNaN(columnIndex)) {
                let cells = rows[columnIndex].cells
                for (let cellIndex in cells) {
                    if (!isNaN(cellIndex)) {
                        sheet.row(Number(columnIndex) + 1)
                            .cell(Number(cellIndex) + 1)
                            .value(cells[cellIndex].text)
                            .style("horizontalAlignment", "center")
                            .style("verticalAlignment", "center")
                            .style("border", true)
                    }
                }
            }

        }
        let M = {}
        merges.forEach(item => {
            M[item] = {
                name: "mergeCell",
                attributes: { ref: item },
                children: []
            }
        })
        sheet._mergeCells = M
        this.generateBlob(newWorkbook)

    },
}