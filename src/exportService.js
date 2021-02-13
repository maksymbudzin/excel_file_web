const xlsx = require('xlsx');
const path = require('path');

const exportExcel = (data, workSheetColumnNames, workSheetName, filePath) => {
    const workBook = xlsx.utils.book_new();
    const workSheetData = [
        workSheetColumnNames,
        ... data
    ];
    const workSheet = xlsx.utils.aoa_to_sheet(workSheetData);
    xlsx.utils.book_append_sheet(workBook, workSheet, workSheetName);
    xlsx.writeFile(workBook, path.resolve(filePath));
}

const exportUsersToExcel = (users, workSheetColumnNames, workSheetName, filePath) => {
    const data = users.map(user => {
        return [user.id, user.text1, user.text2, user.text3, user.stym1, user.stym2, user.stym3, user.stym4, user.stym5,
            user.stym6, user.stym7,user.stym8, user.stym9, user.instr1, user.k, user.l, user.m, user.n, user.o, user.p,
            user.q, user.r, user.s, user.instr2, user.u, user.v, user.w, user.x, user.y, user.z, user.aa, user.ab,
            user.ac, user.instr3, user.ae, user.af, user.ag, user.ah, user.ai, user.aj, user.ak, user.al, user.am,
            user.bool_field];
    });
    exportExcel(data, workSheetColumnNames, workSheetName, filePath);
}

module.exports = exportUsersToExcel;