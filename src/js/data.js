const columns = [
    {
        title: '单位',
        key: 'unit',
        handleSpan: 'UnitSpan',
        rowspan: 2
    },
    {
        title: '时间',
        key: 'date',
        handleSpan: 'UnitSpan',
        rowspan: 2
    },
    {
        title: '单量',
        colspan: 4,
        children: [
            {
                title: 'A单量',
                key: 'numA'
            },
            {
                title: 'B单量',
                key: 'numB'
            },
            {
                title: 'C单量',
                key: 'numC'
            },
            {
                title: 'D单量',
                key: 'numD'
            }
        ]
    }
]

const tableData = [
    {
        unit: '测试单位',
        numA: 1,
        numB: 1,
        numC: 1,
        numD: 1,
        date: 1,
        unitStart: true,
        UnitSpan: 2
    },
    {
        unit: '测试单位',
        numA: 2,
        numB: 2,
        numC: 2,
        numD: 2,
        date: 1,
        UnitSpan: 2
    },
    {
        unit: '合并',
        numA: 2,
        numB: 2,
        numC: 2,
        numD: 2,
        UnitSpan: 2,
        colspanIndexStart: 0,
        colspanNum: 2
    }
]
const title = '表格里显示的标题'
const searchInfo = "表格里显示条件文案"
const excelName = "岛个表格"
const imgBase64 = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' // base64 格式


const params = { columns, tableData, title, searchInfo, excelName, imgBase64 }
export default params