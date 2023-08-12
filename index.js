const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();



const array = [
    {id: 1, name: 'John Doe', age: 35},
    {id: 1, name: 'John Doe', age: 35},
    {id: 1, name: 'John Doe', age: 35},
    {id: 1, name: 'John Doe', age: 35},
]

const worksheet = workbook.addWorksheet('produtos');
worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    {
        header: 'Age', key: 'age', style: { font: { name: 'Arial Black' } } }
]

worksheet.getRow(1).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };

array.forEach((row) => {
    worksheet.addRow(row);
})

worksheet.addRows(array)

workbook.xlsx.writeFile('teste.xlsx').then(() => {
    console.log('Deu bom')
})

// add linha

teste
novo commit
alterar
ai ai
testete
staestgateagdfa
