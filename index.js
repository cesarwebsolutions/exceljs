const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();
// workbook.creator = 'Ng Wai Foong';
// workbook.lastModifiedBy = 'Bot';
// workbook.created = new Date(2021, 8, 30);
// workbook.modified = new Date();
// workbook.lastPrinted = new Date(2021, 7, 27);



const array = [
    { id: 1, name: 'John Doe', age: 35 },
    { id: 1, name: 'John Doe', age: 35 },
    { id: 1, name: 'John Doe', age: 35 },
    { id: 1, name: 'John Doe', age: 35 },
]


const worksheet = workbook.addWorksheet('produtos');
worksheet.addTable({
    name: 'MyTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: {
        theme: 'TableStyleDark3',
        showRowStripes: true,
    },
    columns: [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 32 },
        {
            header: 'Age', key: 'age'
        }
    ],
    row: []
});
const table = worksheet.getTable('MyTable');
const column = table.getColumn(1);
column.name = 'Code';
column.filterButton = true;
column.style = { font: { bold: true, name: 'Comic Sans MS' } };
column.totalsRowLabel = 'Totals';
column.totalsRowFunction = 'custom';
column.totalsRowFormula = 'ROW()';
column.totalsRowResult = 10;
column.totalsRowLabel = 'Totals';


worksheet.commit()

table.getRow(1).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };

array.forEach((row) => {
    table.addRow(row);
})

table.addRows(array)

workbook.xlsx.writeFile('teste.xlsx').then(() => {
    console.log('Deu bom')
})
