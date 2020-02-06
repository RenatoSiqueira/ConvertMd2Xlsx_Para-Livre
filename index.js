const lineReader = require('line-reader')
var xl = require('excel4node')

var wb = new xl.Workbook()
var ws = wb.addWorksheet('Comunidades de tecnologia PA')

var titleStyle = wb.createStyle({ font: { size: 12, bold: true } })

let Row = 1
let Column = 1

lineReader.eachLine('README.md', line => {
    // Seções
    if (line.match(/#.+/gm)) {
        ws
            .cell(Row, Column, Row, Column + 6, true)
            .string(line.slice(2).toUpperCase())
            .style(titleStyle)
        Row++
    }

    // Comunidade
    if (line.match(/-.\*\*.+/gm)) {

        const lineClean = line.replace(/:t.+t:$/gm, '')

        if (lineClean.match(/:/gm)) {
            lineClean.replace(/(-.\*\*.+):(.+)/gm, (_, key, value) => {
                // Nome Comunidade
                ws.cell(Row, Column).string(key.replace(/\*/gm, ''))
                Row++
                // Descrição Comunidade
                ws.cell(Row, Column).string(value.trim())
                Row++
            })
        } else {
            // Nome Comunidade
            ws.cell(Row, Column).string(lineClean.replace(/\*/gm, ''))
            Row++
        }
    }

    // Dados da Comunidade
    if (line.match(/-.\[.+/gm)) {
        ws.cell(Row, Column + 1).string(line)
        Row++
    }

}, () => {
    wb.write('Excel.xlsx')
    console.log('End')
})
