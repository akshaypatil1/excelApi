const xl = require('excel4node');

let sampleJSON = [{ header: 'Some header here' }]
for (let index = 1; index <= 500; index++) {
    sampleJSON.push(
        { name: `Name_${index}`, age: `age_${index}`, address: `address_${index}`, phone: `Phone_${index}` }
    )
}
async function get(req, res, next) {
    try {
        // Create a new instance of a Workbook class
        let wb = new xl.Workbook();
        // Add Worksheets to the workbook
        let ws = wb.addWorksheet('Report');
        let hederStyle = wb.createStyle({
            font: {
                bold: true,
                color: '00FF00',
            },
            alignment: {
                wrapText: true,
                horizontal: 'center',
            },
        });
        sampleJSON.forEach((obj, i) => {
            let row = i + 1;
            if (obj.hasOwnProperty('header')) {
                ws.cell(row, 1, row, 4, true).string(obj.header).style(hederStyle);
            } else {
                let colKeys = Object.keys(obj);
                colKeys.forEach((key, colIndex) => {
                    let col = colIndex + 1;
                    ws.cell(row, col).string(obj[key]);
                });
            }
        });
        wb.write('Excel.xlsx', res);
    } catch (e) {
        next(e);
    }
}
module.exports.get = get;
