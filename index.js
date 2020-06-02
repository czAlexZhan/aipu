const xlsx = require('node-xlsx');
const fs = require('fs');

// 销售单
const file1 = xlsx.parse(`./files/1.xlsx`);
// 出货单
const file2 = xlsx.parse(`./files/2.xlsx`);

const shipper = file1[0].data;
const receivable = file2[0].data;
// 销售单title
const shipperTitle = shipper[1];
// 出货单title
const receivableTitle = receivable[1];

// 销售单 合同编号index
const shipperDan = shipperTitle.indexOf('合同编号');
// 销售单 料号index
const shipperLiao = shipperTitle.indexOf('料号');

// 出货单 合同编号
const receivableDan = receivableTitle.indexOf('合同编号');
// 出货单 料号
const receivableLiao = receivableTitle.indexOf('料号');

const shipperMap = new Map();
const receivableMap = new Map();

for(let i=2,len=shipper.length;i<len;i++) {
  const _d = shipper[i];
  const bianHao = _d[shipperDan];
  const liaoHao = _d[shipperLiao];
  if(!bianHao) continue;
  if(shipperMap.has(bianHao)) {
    const map = shipperMap.get(bianHao);
    map.set(liaoHao, _d);
  }else {
    const map = new Map();
    map.set(liaoHao, _d);
    shipperMap.set(bianHao, map);
  }
}

for(let i=2,len=receivable.length;i<len;i++) {
  const _d = receivable[i];
  const bianHao = _d[receivableDan];
  const liaoHao = _d[receivableLiao];
  if(receivableMap.has(bianHao)) {
    const map = receivableMap.get(bianHao);
    map.set(liaoHao, _d);
  }else {
    const map = new Map();
    map.set(liaoHao, _d);
    receivableMap.set(bianHao, map);
  }
}


const outData = [];
// 单号
const danHao = shipperTitle.indexOf('单号');
// 合同编号
const heTongBianHao = shipperDan;
// 项目名称
const xiangMu = shipperTitle.indexOf('项目名称');
// 料号
const liaoHao = shipperTitle.indexOf('料号');
// 料品名称
const liaoPin = shipperTitle.indexOf('料品名称');
// 最终价
const zuiZhongJia = receivableTitle.indexOf('最终价');
// 累计出货计划数量
const dingHuo = shipperTitle.indexOf('累计出货计划数量');
// 价税合计
const jiaShui = shipperTitle.indexOf('价税合计');
// 已出货数量
const yiChuHuo = receivableTitle.indexOf('数量');
// 收货地址
const address = receivableTitle.indexOf('收货地址');
// 遍历销售单

const titleRow = ['单号', '合同编号', '项目名称', '料号', '料品名称', '最终价', '累计出货计划数量', '价税合计', '数量', '收货地址'];
outData.push(titleRow);
shipperMap.forEach((val, key) => {
  let chuHuoRow = null;
  let huoWuMap = null;
  try {
    // 出货单相同的合同编号
    huoWuMap = receivableMap.get(key);
    if(val.size !== 0 && key) {
      val.forEach((shipperRow, key1) => {
        // 货物对应的料号
        chuHuoRow = (huoWuMap && huoWuMap.get(key1))?huoWuMap.get(key1): new Array(receivableTitle.length).fill('空');
        const data = [shipperRow[danHao], shipperRow[heTongBianHao], shipperRow[xiangMu], shipperRow[liaoHao],
          shipperRow[liaoPin], chuHuoRow[zuiZhongJia], shipperRow[dingHuo], shipperRow[jiaShui], chuHuoRow[yiChuHuo], chuHuoRow[address]];
        outData.push(data);
      })
    }
  }catch (e) {
      console.log(e)
  }
});

const buffer = xlsx.build([{name: '对账单', data: outData}]);
fs.writeFileSync('./files/3.xlsx', buffer);


