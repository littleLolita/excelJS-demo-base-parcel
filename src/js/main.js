
import Excel from 'exceljs'
import * as FileSaver from 'file-saver'
import params from './data'

/**
 * @param {*} columns //表头
 * @param {*} tableData //表格数据
 * @param {*} title //标题
 * @param {*} searchInfo //搜素条件
 * @param {*} excelName //导出名称
 * @param {*} imgBase64 //导出图片，base64格式
 */
function exportImgExcelTable (params) {
  const { columns, tableData, title, searchInfo, excelName, imgBase64 } = params
  const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'
  // 创建工作簿
  let workbook = new Excel.Workbook()
  // 设置工作簿的属性
  workbook.creator = 'Me'
  workbook.lastModifiedBy = 'Her'
  workbook.created = new Date(1985, 8, 30)
  workbook.modified = new Date()
  workbook.lastPrinted = new Date()
  // 工作簿视图，工作簿视图控制在查看工作簿时 Excel 将打开多少个单独的窗口
  workbook.views = [
    {
      x: 0,
      y: 0,
      width: 1000,
      height: 2000,
      firstSheet: 0,
      activeTab: 1,
      visibility: 'visible'
    }
  ]
  let worksheet = workbook.addWorksheet('sheet1')

  let baseRow = 1

  // 是否有图片
  if (imgBase64) {
    // 通过 base64  将图像添加到工作簿
    const imageId1 = workbook.addImage({
      base64: imgBase64,
      extension: 'jpg'
    })
    //  在一定范围内添加图片
    worksheet.addImage(
      imageId1,
      `A1:C12`
    )
    baseRow = 6 // 根据图片比例定
  }

  // 计算表格列数
  let len = 0
  const ishasChildren = columns.filter(item => item.children)
  if (ishasChildren.length > 0) {
    columns.map(item => {
      if (item.children) {
        len += item.colspan
      } else {
        len += 1
      }
    })
  } else {
    len = columns.length
  }

  // 是否有标题
  if (title) {
    worksheet.insertRow(baseRow, [title])
    worksheet.mergeCells(`A${baseRow}:${String.fromCharCode(65 + len - 1)}${baseRow}`)
    worksheet.getRow(baseRow).font = {
      size: 16,
      bold: true
    }
    worksheet.getRow(baseRow).alignment = { vertical: 'middle', horizontal: 'center' }
    worksheet.getRow(baseRow).height = 30
    baseRow += 1
  }
  // 是否有搜错条件
  if (searchInfo) {
    worksheet.insertRow(baseRow, [searchInfo])
    worksheet.mergeCells(`A${baseRow}:${String.fromCharCode(65 + len - 1)}${baseRow}`)
    worksheet.getRow(baseRow).alignment = { vertical: 'middle' }
    worksheet.getRow(baseRow).height = 25
    baseRow += 1
  }

  // 表头处理: 单行表头或多行表头合并
  const hasChildren = columns.filter(item => item.children)
  let headArray = []
  let listHearder = []

  if (hasChildren.length > 0) {
    headArray = [[], []]
    listHearder = []
    columns.map((item, index) => {
      if (item.children && item.colspan) {
        headArray[0].push(item.title)
        for (let i = 0; i <= item.colspan - 2; i++) {
          headArray[0].push('')
        }
        item.children.map((ele, ind) => {
          headArray[1].push(ele.title)
          listHearder.push(ele.key)
        })
      } else {
        headArray[0].push(item.title)
        headArray[1].push('')
        listHearder.push(item.key)
      }
    })

    worksheet.insertRow(baseRow, headArray[0])
    worksheet.insertRow(baseRow + 1, headArray[1])
    let activeMerge = 0
    columns.map((item, index) => {
      if (item.children && item.colspan) {
        const mergeStartIndex = String.fromCharCode(65 + index + activeMerge)
        const mergeEndIndex = String.fromCharCode(65 + index + activeMerge + item.colspan - 1)
        const mergeCode = `${mergeStartIndex}${baseRow}:${mergeEndIndex}${baseRow}`
        worksheet.mergeCells(mergeCode)
        activeMerge += item.colspan - 1
      } else {
        const columnsIndex = String.fromCharCode(65 + index)
        const mergeCode = `${columnsIndex}${baseRow}:${columnsIndex}${baseRow + 1}`
        worksheet.mergeCells(mergeCode)
      }
    })

    worksheet.getRow(baseRow)._cells.forEach((item, index) => {
      worksheet.getCell(item._address).alignment = { vertical: 'middle', horizontal: 'center' }
      worksheet.getColumn(index + 1).width = 20
    })
    worksheet.getRow(baseRow).font = { bold: true }
    worksheet.getRow(baseRow + 1)._cells.forEach((item, index) => {
      worksheet.getCell(item._address).alignment = { vertical: 'middle', horizontal: 'center' }
      worksheet.getColumn(index + 1).width = 20
    })
    worksheet.getRow(baseRow + 1).font = { bold: true }
    baseRow += 2
  } else {
    headArray = columns.map((item) => { return item.title })
    listHearder = columns.map((item) => { return item.key })
    worksheet.insertRow(baseRow, headArray)

    worksheet.getRow(baseRow)._cells.forEach((item, index) => {
      worksheet.getCell(item._address).alignment = { vertical: 'middle', horizontal: 'center' }
      worksheet.getColumn(index + 1).width = 20
    })

    worksheet.getRow(baseRow).font = { bold: true }
    baseRow += 1
  }
  // 插入表格数据
  const tableDataList = tableData
  let array = []
  tableData.forEach((item, index) => {
    array.push(
      worksheet.insertRow(index + baseRow, listHearder.map((ite) => item[ite]))
    )
  })

  // 合并单元格处理
  tableDataList.forEach((item, index) => {
    if (item.colspanNum) {
      const mergeStartIndex = String.fromCharCode(65 + item.colspanIndexStart)
      const mergeEndIndex = String.fromCharCode(65 + item.colspanIndexStart + item.colspanNum - 1)
      const mergeCode = `${mergeStartIndex}${index + baseRow}:${mergeEndIndex}${index + baseRow}`

      if (mergeStartIndex !== mergeEndIndex) {
        worksheet.mergeCells(mergeCode)
      }
    }

    columns.map((col, colIndex) => {
      if (col.needHandle) {
        const key = col.handleSpan
        if (item[key] && item[`${key}Start`]) {
          const columnsIndex = String.fromCharCode(65 + colIndex)
          const mergeCode = `${columnsIndex}${index + baseRow}:${columnsIndex}${index + baseRow + item[key] - 1}`
          if (item[key] - 1 > 0) {
            worksheet.mergeCells(mergeCode)
          }
        }
      }
    })
  })
  // 样式居中
  for (const item of array) {
    item.eachCell((cell) => {
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    })
  }
  // 导出表格数据
  workbook.xlsx.writeBuffer().then((data) => {
    const blob = new Blob([data], { type: EXCEL_TYPE })
    FileSaver.saveAs(blob, `${excelName}.xlsx`)
  })
}

 
setTimeout(() => {
    exportImgExcelTable(params)
}, 1000)
