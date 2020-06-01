const xlsx = require('node-xlsx');
// 出货单
const file1 = xlsx.parse(`./files/1.xlsx`);
// 应收单
const file2 = xlsx.parse(`./files/2.xlsx`);

const shipper = file1[0].data;
const receivable = file2[0].data;
// 出货单title
const shipperTitle = shipper[1];
// 应收单title
const receivableTitle = receivable[1];

// 出货单 单据编号index
const shipperDan = shipperTitle.indexOf('单据编号');
// 出货单 料号index
const shipperLiao = shipperTitle.indexOf('料号');

// 应收单 来源单据编号
const receivableDan = receivableTitle.indexOf('来源单据编号');
// 应收单 料号
const receivableLiao = receivableTitle.indexOf('料号');
