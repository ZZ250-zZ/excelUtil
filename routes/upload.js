// 相当于 controller，定义接口路径

const fs = require('fs');
const express = require('express');
const multer = require('multer');
const path = require('path');
const router = express.Router();
const bodyParser = require('body-parser')
const {Workbook, ValueType} = require('exceljs');
const Decimal = require('decimal.js');
const AdmZip = require('adm-zip');

const SimpleCache = require('../util/simpleCache');
const excelCache = new SimpleCache();

//创建 application/json 解析
const jsonParser = bodyParser.json();
//创建 application/x-www-form-urlencoded
const formParser = bodyParser.urlencoded({extended: false});

// 佣金默认值，匹配不到佣金时的填充
var defaultYj = 'NULL';

// 总表商务到店铺的映射
function sw2dp(zc, sw) {
  if (zc === '吴鑫') {
    return sw.slice(0, -1)
  } else {
    return sw.slice(0, -2)
  }
}

function mkdirsSync(dirname) {
  if (fs.existsSync(dirname)) {
    return true;
  } else {
    if (mkdirsSync(path.dirname(dirname))) {
      fs.mkdirSync(dirname);
      return true;
    }
  }
}

// 更改大文件的存储路径
var createFolder = function (folder) {
  try {
    fs.accessSync(folder);
  } catch (e) {
    mkdirsSync(folder);
  }
};
var uploadFolder = '/tmp/upload';// 设定存储目录为 /tmp/upload（Linux 的内存文件系统）。现在改成了自己写内存
// createFolder(uploadFolder);

// 磁盘存贮
var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    // 清空上传文件夹
    if (!fs.existsSync(uploadFolder)) {
      fs.mkdirSync(uploadFolder, { recursive: true });
    } else {
      fs.readdir(uploadFolder, (err, files) => {
        if (err) throw err;

        for (const file of files) {
          fs.unlink(path.join(uploadFolder, file), (err) => {
            if (err) throw err;
          });
        }
      });
    }

    cb(null, uploadFolder); // 他会放在当前目录下的 /upload 文件夹下（没有该文件夹，就新建一个）
  },
  filename: function (req, file, cb) { // 在这里设定文件名
    cb(null, file.originalname);
  }
})

// 文件写磁盘
var multerDisk = multer({'storage': storage})

// 文件存到内存
const multerMem = multer({'storage': multer.memoryStorage()})


// 跳转到 index.html
router.get('/', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/index.html'))
})

// 跳转到 yongjinjisuan.html
router.get('/yongjinjisuan', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/yongjinjisuan.html'))
})

// 跳转到 biaogechaifen.html
router.get('/biaogechaifen', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/biaogechaifen.html'))
})

// 上传文件
router.post('/yjjsAction', multerMem.fields([{name: 'bdFile', maxCount: 1}, {name: 'yjFile', maxCount: 1}]),
  async function (req, res) {
    console.dir(req.files)

    if (!req.files || Object.keys(req.files).length === 0) {
      res.status(400).send('请选择要上传的文件！');
      return;
    }


    // 日期
    var date = req.body.date

    // 佣金
    let cache = new Map();
    const fileYJ = req.files.yjFile[0].buffer;
    let workbookYJ = await new Workbook().xlsx.load(fileYJ);

    workbookYJ.eachSheet((sheet, index1) => {
      sheet.eachRow((row, rowIdx) => {
        let rowData = [];
        row.eachCell({includeEmpty: true}, function (cell, colNumber) {
          rowData.push(cell.value);
        });

        let key = rowData[0]
        const val = {min: rowData[1], max: rowData[2], yj: rowData[3]}
        // const val = {min: new Decimal(rowData[1]), max: new Decimal(rowData[2]), yj: new Decimal(rowData[3])}
        var newVar = cache.get(key);
        if (!newVar) {
          cache.set(key, [val])
        } else {
          newVar.push(val)
        }
      })
    })
    workbookYJ = null
    // console.log(cache)

    // 补单文件
    const file = req.files.bdFile[0].buffer;
    let workbook2 = await new Workbook().xlsx.load(file);

    var columns = [
      {header: '商务', key: 'sw', width: 15},
      {header: '日期', key: 'rq', width: 10},
      {header: '店铺名', key: 'dpm', width: 30},
      {header: '会员名', key: 'hym', width: 25},
      {header: '订单号', key: 'ddh', width: 25},
      {header: '价格', key: 'jg', width: 10},
      {header: '佣金', key: 'yj', width: 10},
      {header: '主持佣金', key: 'zcyj', width: 10},
      {header: '主持', key: 'zc', width: 10}
    ]
    var allData = []
    workbook2.eachSheet((sheet, index1) => {
      // console.log('工作表' + index1);
      sheet.eachRow((row, rowIdx) => {
        let rowData = [];
        row.eachCell({includeEmpty: true}, function (cell, colNumber) {
          rowData.push(cell.value);
        });

        // 输出当前行的内容
        // console.log(rowData)
        if (rowIdx === 1) {
          rowData.splice(6, 0, '佣金')
        }
        else {
          // console.log(rowData)
          const jg = rowData[5];
          const valArr = cache.get(rowData[0])
          if (valArr) {
            const found = valArr.find((element) => jg >= element.min && jg <= element.max);
            if (found) {
              rowData.splice(6, 0, found.yj)
            } else {
              rowData.splice(6, 0, defaultYj)
            }
          } else {
            rowData.splice(6, 0, defaultYj)
          }
          allData.push({
            sw: rowData[0],
            rq: date,
            dpm: rowData[2],
            hym: rowData[3],
            ddh: rowData[4],
            jg: jg.toString(),
            yj: rowData[6].toString(),
            zcyj: rowData[7].toString(),
            zc: rowData[8],
          })
        }
      })
    })
    workbook2 = null

    // 写文件
    const downWB = new Workbook();
    const downWS = downWB.addWorksheet('Sheet1');
    downWS.columns = columns
    downWS.addRows(allData);
    downWS.eachRow((row, rowIndex) => {
      row.eachCell(cell => {
        if (rowIndex === 1) {
          cell.style = {
            font: {size: 11, bold: true},
            alignment: {vertical: 'middle', horizontal: 'center'},
            border: {
              top: {style: 'thin', color: {argb: 'thin'}},
              left: {style: 'thin', color: {argb: 'thin'}},
              bottom: {style: 'thin', color: {argb: 'thin'}},
              right: {style: 'thin', color: {argb: 'thin'}}
            }
          }
        } else {
          cell.style = {
            font: {size: 11, bold: false,},
            alignment: {vertical: 'middle', horizontal: 'center'},
            border: {
              top: {style: 'thin', color: {argb: 'thin'}},
              left: {style: 'thin', color: {argb: 'thin'}},
              bottom: {style: 'thin', color: {argb: 'thin'}},
              right: {style: 'thin', color: {argb: 'thin'}}
            }
          }
        }
      })
    })

    const fileName = 'temp_' + date + '_lx.xlsx'
    const buffer = await downWB.xlsx.writeBuffer();
    const mapKey = encodeURIComponent(fileName);
    excelCache.set(mapKey, buffer);
    // 释放引用
    req.files = null;


    const list = [{name: fileName, path: mapKey}];
    res.render('filelist.ejs', {list: list, title: '佣金计算'})
  });


// 下载文件
router.get('/downloadBuffer', function (req, res) {
  // query 参数，接收过来已经被 decodeURIComponent 了
  var fileName = req.query.path;
  var cacheKey = encodeURIComponent(fileName);

  // console.log(new Date().toLocaleString(), '下载文件：' + fileName)
  var fuleBuffer = excelCache.get(cacheKey);
  if (!fuleBuffer) {
    res.status(400).send('文件已删除，请重新执行上一步操作！');
    return;
  }

  res.set('Content-Disposition', `attachment; filename*=utf-8''${cacheKey}`);
  res.set('Content-Type', 'application/octet-stream');
  res.send(fuleBuffer)
})


router.get('/test', function (req, res) {
  const list = [{name: 'temp_3.5_lx.xlsx', path: '/tmp/upload/temp_3.5_lx.xlsx'}];
  res.render('filelist.ejs', {list: list, title: '佣金计算'})
})

// 设置excel样式
function formatExcel(downWS) {
  downWS.eachRow((row, rowIndex) => {
    row.eachCell(cell => {
      if (rowIndex === 1) {
        cell.style = {
          font: {
            size: 11,
            bold: true,
            color: {argb: '000000'}
          },
          alignment: {vertical: 'middle', horizontal: 'center'},
          border: {
            top: { style: 'thin', color: { argb: '000000' } },
            left: { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            right: { style: 'thin', color: { argb: '000000' } }
          }
        }
      } else {
        cell.style = {
          font: {
            size: 11,
            bold: false,
          },
          alignment: {vertical: 'middle', horizontal: 'center'},
          border: {
            top: { style: 'thin', color: { argb: '000000' } },
            left: { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            right: { style: 'thin', color: { argb: '000000' } }
          }
        }
      }
    })
  })
}

// 设置excel样式：店铺资金统计
function dpzjtjFormatExcel(downWS) {
  downWS.eachRow((row, rowIndex) => {
    row.eachCell((cell, colNumber) => {
      if (rowIndex === 1) {
        cell.style = {
          font: {size: 11, bold: true, color: {argb: '000000'}},
          alignment: {vertical: 'middle', horizontal: 'center'},
          border: {top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: '000000' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } }}
        }
      } else {
        cell.style = {
          font: {size: 11, bold: false,},
          alignment: {vertical: 'middle', horizontal: 'center'},
          border: {top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: '000000' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: '000000' } }}
        }
        if (colNumber === 6) cell.style.alignment.horizontal = 'left'
      }
    })
  })
}


// 带佣金的补单表头
const yjColumns = [
  {header: '商务', key: 'sw', width: 15},
  {header: '日期', key: 'rq', width: 10},
  {header: '店铺名', key: 'dpm', width: 30},
  {header: '会员名', key: 'hym', width: 30},
  {header: '订单号', key: 'ddh', width: 25},
  {header: '价格', key: 'jg', width: 20},
  {header: '佣金', key: 'yj', width: 20},
  {header: '主持佣金', key: 'zcyj', width: 20},
  {header: '主持', key: 'zc', width: 10}
]
// 主持表头：【导出】3.4博.xlsx
const zcColumns = [
  {header: '商务', key: 'sw', width: 15},
  {header: '日期', key: 'rq', width: 10},
  {header: '店铺名', key: 'dpm', width: 30},
  {header: '会员名', key: 'hym', width: 30},
  {header: '订单号', key: 'ddh', width: 25},
  {header: '价格', key: 'jg', width: 20},
  {header: '佣金', key: 'yj', width: 20},
  {header: '主持佣金', key: 'zcyj', width: 20},
  {header: '主持', key: 'zc', width: 10}
]
// 店铺表头：./【导出】3.4/3.4 勤勉 3单 630元.xlsx
const dpColumns = [
  {header: '日期', key: 'rq', width: 10},
  {header: '店铺名', key: 'dpm', width: 30},
  {header: '会员名', key: 'hym', width: 30},
  {header: '订单号', key: 'ddh', width: 25},
  {header: '价格', key: 'jg', width: 20},
  {header: '佣金', key: 'yj', width: 20}
]
// 店铺资金统计：【导出】店铺资金统计表.xlsx
const dpzjtjColumns = [
  {header: '日期', key: 'rq', width: 10},
  {header: '店铺名', key: 'dpm', width: 30},
  {header: '今日小计', key: 'jrxj', width: 20},
  {header: '昨日差额', key: 'zrce', width: 20},
  {header: '总计', key: 'zj', width: 20},
  {header: '公式', key: 'gs', width: 50},
  {header: 'Excel路径', key: 'lj', width: 10}
]

// 凌星根目录：         ./凌星
const lxRootPath = '凌星'


// 上传文件
router.post('/parse', multerMem.fields([{name: 'bdFile', maxCount: 1}]),
  async function (req, res) {
    if (!req.files || Object.keys(req.files).length === 0) {
      res.status(400).send('请选择要上传的文件！');
      return;
    }

    // 获取日期
    const originalname = req.files.bdFile[0].originalname
    const date = originalname.substring(5, originalname.indexOf('_', 5));
    // 前缀: 【导出】3.4
    const prefix = '【导出】' + date;

    // ./凌星/【导出】3.4凌星.xlsx
    let lxList = []
    // ./凌星/【导出】店铺资金统计表.xlsx
    let lxDpzjtjList = []
    //  lx店铺明细 目录：   ./【导出】3.4
    const lxDpmxDir = lxRootPath + '/' + prefix;


    // 主持分组
    let cache = new Map();

    // 解析替换过 null 的 temp_3.4_lx.xlsx
    const file = req.files.bdFile[0].buffer;
    let workbook2 = await new Workbook().xlsx.load(file);

    workbook2.eachSheet((sheet, index1) => {
      sheet.eachRow((row, rowIdx) => {
        let rowData = [];
        row.eachCell({includeEmpty: true}, function (cell, colNumber) {
          rowData.push(cell.value);
        });

        if (rowIdx !== 1) {
          let key = rowData[8]
          // 完整的补单信息
          const bdObj = {
            sw: rowData[0], rq: rowData[1], dpm: rowData[2], hym: rowData[3], ddh: rowData[4],
            jg: rowData[5], yj: rowData[6], zcyj: rowData[7], zc: rowData[8]
          }
          if (bdObj.yj === defaultYj) {
            res.status(400).send(`无效的佣金：${bdObj.dpm}`);
            return;
          }
          var mapValue = cache.get(key);
          if (mapValue) {
            mapValue.push(bdObj)
          } else {
            cache.set(key, [bdObj])
          }

          lxList.push(bdObj)
        }
      })
    })


    // 创建 主持zip
    const lxZipFile = new AdmZip();

    // 遍历主持分组，写文件
    // 用来渲染页面的文件数组
    const list = [];
    for (let item of cache) {
      // 主持人：吴鑫
      const zc = item[0]
      const value = item[1]

      // 主持 根目录   ./博
      const zcRootPath = zc;
      // 创建 主持zip
      const zcZipFile = new AdmZip();


      const downWB = new Workbook();
      const downWS = downWB.addWorksheet('Sheet1');
      downWS.columns = yjColumns
      downWS.addRows(value);
      formatExcel(downWS);
      // 文件名：     ./博【导出】3.4博.xlsx
      const tempFileName = prefix + zc + '.xlsx'
      const filePath = zcRootPath + '/' + tempFileName;
      // console.log(filePath);
      zcZipFile.addFile(filePath, await downWB.xlsx.writeBuffer())

      //  主持的店铺明细目录：   ./博/【导出】3.4
      const zcDpmxDir = zcRootPath + '/' + prefix;
      let dianpuGroup = value.reduce((groups, item) => {
        let groupName = sw2dp(item.zc, item.sw);
        if (!groups[groupName]) {
          groups[groupName] = [];
        }
        groups[groupName].push(item);
        return groups;
      }, {});

      // 【导出】店铺资金统计表.xlsx
      let dpzjtjList = []
      // 生成主持下的店铺汇总表
      for (const dpName in dianpuGroup) {
        const tempList = dianpuGroup[dpName]
        let sum = tempList.reduce((total, item) => {
          return total.plus(new Decimal(item.jg)).plus(new Decimal(item.yj));
        }, new Decimal(0));
        let totalPrice = sum.toFixed(2);
        // 店铺明细文件：      3.4 勤勉 3单 630元.xlsx
        const dpMxFileName = date + ' ' + dpName + ' ' + tempList.length +'单 '+ totalPrice + '元.xlsx'
        tempList.push({})
        tempList.push({jg: '合计', yj: totalPrice})

        // TODO 1，主持目录的店铺明细:             ./博/【导出】3.4/3.4 饭饭 26单 6607元.xlsx
        const dpMxPath = zcDpmxDir + '/' + dpMxFileName;
        const dpMxWB = new Workbook();
        const dpMxWS = dpMxWB.addWorksheet('Sheet1');
        dpMxWS.columns = dpColumns
        dpMxWS.addRows(tempList);
        formatExcel(dpMxWS);
        zcZipFile.addFile(dpMxPath, await dpMxWB.xlsx.writeBuffer())

        // TODO 2，凌星目录的店铺明细:            ./凌星/【导出】3.4/3.4 饭饭 26单 6607元.xlsx
        const lxDpMxPath = lxDpmxDir + '/' + dpMxFileName;
        const lxDpMxWB = new Workbook();
        const lxDpMxWS = lxDpMxWB.addWorksheet('Sheet1');
        lxDpMxWS.columns = dpColumns
        lxDpMxWS.addRows(tempList);
        formatExcel(lxDpMxWS);
        // console.log(lxDpMxPath);
        lxZipFile.addFile(lxDpMxPath, await lxDpMxWB.xlsx.writeBuffer())


        let zrceTemp = new Decimal(0)
        let zjTemp = sum.plus(zrceTemp).toFixed(2)
        let index = dpzjtjList.length + 2
        const zjtjObj = {
          rq: date, dpm: dpName, jrxj: totalPrice, zrce: '0', lj: '',
          zj: {formula: `=C${index}+D${index}`},
          gs: {formula: `=A${index}&"  今日合计"&C${index}&"， 昨日差额"&D${index}&"，  合计差额"&E${index}`}
        }
        dpzjtjList.push(zjtjObj)

        // 凌星的 店铺资金统计
        lxDpzjtjList.push(zjtjObj)
      }

      const dpzjWB = new Workbook();
      const dpzjWS = dpzjWB.addWorksheet('Sheet1');
      dpzjWS.columns = dpzjtjColumns
      dpzjWS.addRows(dpzjtjList);
      dpzjtjFormatExcel(dpzjWS);
      const dpzjPath = zcRootPath + '/【导出】店铺资金统计表.xlsx';
      zcZipFile.addFile(dpzjPath, await dpzjWB.xlsx.writeBuffer())



      const zipName = '_' + zc + '.zip'
      const zcCacheKey = encodeURIComponent(zipName);
      excelCache.set(zcCacheKey, zcZipFile.toBuffer());
      list.push({name: zipName, path: zcCacheKey})
    }


    // 保存       ./凌星/【导出】3.4凌星.xlsx
    const lxWB = new Workbook();
    const lxWS = lxWB.addWorksheet('Sheet1');
    lxWS.columns = yjColumns
    lxWS.addRows(lxList);
    formatExcel(lxWS);
    // 文件名：     【导出】3.4凌星.xlsx
    const lxFileName = prefix + '凌星.xlsx'
    // 凌星根目录：         ./凌星/【导出】3.4凌星.xlsx
    const lxPath = lxRootPath + '/' + lxFileName;
    console.log(lxPath);
    lxZipFile.addFile(lxPath, await lxWB.xlsx.writeBuffer())

    // 保存       ./凌星/【导出】店铺资金统计表.xlsx
    const lxDpzjWB = new Workbook();
    const lxDpzjWS = lxDpzjWB.addWorksheet('Sheet1');
    lxDpzjWS.columns = dpzjtjColumns
    lxDpzjWS.addRows(lxDpzjtjList);
    dpzjtjFormatExcel(lxDpzjWS);
    const lxDpzjPath = lxRootPath + '/【导出】店铺资金统计表.xlsx';
    console.log(lxDpzjPath);
    lxZipFile.addFile(lxDpzjPath, await lxDpzjWB.xlsx.writeBuffer())

    // 创建         _凌星.zip
    // const lxZipFile = new AdmZip();
    // lxZipFile.addLocalFolder(lxRootPath, '凌星');
    const lxFilaName = '_凌星.zip'

    const lxCacheKey = encodeURIComponent(lxFilaName);
    excelCache.set(lxCacheKey, lxZipFile.toBuffer());
    list.push({name: lxFilaName, path: lxCacheKey})

    workbook2 = null
    req.files = null
    res.render('filelist.ejs', {list: list, title: '表格拆分'})
  });


// 上传页面
router.get('/index', (req, res) => {
  console.log(__dirname)
  res.sendFile(path.join(__dirname, '../demo/upload.html'))
})

// 列表页面
router.get('/filelist', function (req, res) {
  res.sendFile(path.join(__dirname, '../demo/filelist.html'))
})

// 上传文件
router.post('/upload', multerDisk.array('file'), function (req, res) {
  console.dir(req.files)

  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).send('请选择要上传的文件！');
    return;
  }

  // res.send('Success.');
  // 重定向到列表页
  res.redirect('/filelist')
});

// 下载文件
router.get('/download', function (req, res) {
  var filePath = req.query.path;
  console.log('下载文件：' + filePath)
  filePath = path.join(__dirname, '../' + filePath);
  res.attachment(filePath)
  res.sendFile(filePath)
})

// 删除文件
router.post('/delete', jsonParser, function (req, res, next) {
  var filePath = req.body.filePath;
  console.log('删除文件：' + filePath)

  try {
    fs.unlinkSync(filePath)
    // 重定向到列表页
    res.send('删除成功')
  } catch (error) {
    res.send('删除失败')
  }

})


// 获取文件列表
router.get('/getFileList', function (req, res, next) {
  var filelist = getFileList(uploadFolder)
  res.send(filelist)
})

function getFileList(path) {
  var filelist = [];
  readFile(path, filelist);

  return filelist;
}


function readFile(path, filelist) {
  var files = fs.readdirSync(path);
  files.forEach(walk);

  function walk(file) {
    var state = fs.statSync(path + '/' + file)
    if (state.isDirectory()) {
      readFile(path + '/' + file, filelist)
    } else {
      var obj = new Object;
      obj.size = state.size;
      obj.name = file;
      obj.path = path + '/' + file;
      filelist.push(obj);
    }
  }
}

module.exports = router;
