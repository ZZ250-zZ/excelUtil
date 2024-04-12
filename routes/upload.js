const fs = require('fs');
const express = require('express');
const multer = require('multer');
const path = require('path');
const router = express.Router();
const bodyParser = require('body-parser')
const {Workbook, ValueType} = require('exceljs');
const Decimal = require('decimal.js');
const AdmZip = require('adm-zip');

//创建application/json解析
var jsonParser = bodyParser.json();

//创建application/x-www-form-urlencoded
var urlencodedParser = bodyParser.urlencoded({extended: false});

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

var uploadFolder = '/tmp/upload';// 设定存储文件夹为当前目录下的 /upload 文件夹
createFolder(uploadFolder);
// 磁盘存贮
var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadFolder); // 他会放在当前目录下的 /upload 文件夹下（没有该文件夹，就新建一个）
  },
  filename: function (req, file, cb) { // 在这里设定文件名
    cb(null, file.originalname);
  }
})

var upload = multer({storage: storage})


// index
router.get('/', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/index.html'))
})

// 佣金计算
router.get('/yongjinjisuan', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/yongjinjisuan.html'))
})

// 表格拆分
router.get('/biaogechaifen', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/biaogechaifen.html'))
})

// 上传文件
router.post('/upload2', upload.fields([{name: 'bdFile', maxCount: 1}, {name: 'yjFile', maxCount: 1}]),
  async function (req, res) {
    console.dir(req)
    console.dir(req.files)

    if (!req.files || Object.keys(req.files).length === 0) {
      res.status(400).send('请选择要上传的文件！');
      return;
    }


    // 日期
    var date = req.body.date

    // 佣金
    let cache = new Map();
    const fileYJ = req.files.yjFile[0].path;
    const workbookYJ = await new Workbook().xlsx.readFile(fileYJ);

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
    console.log(cache)

    // budan
    const file = req.files.bdFile[0].path;
    const workbook = new Workbook();
    const workbook2 = await workbook.xlsx.readFile(file);

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
      console.log('工作表' + index1);
      sheet.eachRow((row, rowIdx) => {
        let rowData = [];
        row.eachCell({includeEmpty: true}, function (cell, colNumber) {
          rowData.push(cell.value);
        });

        // 输出当前行的内容

        if (rowIdx === 1) {
          rowData.splice(6, 0, '佣金')
        } else {
          console.log(rowData)
          const jg = rowData[5];
          const valArr = cache.get(rowData[0])
          if (valArr) {
            const found = valArr.find((element) => jg >= element.min && jg <= element.max);
            if (found) {
              rowData.splice(6, 0, found.yj)
            } else {
              rowData.splice(6, 0, 'NULL')
            }
          } else {
            rowData.splice(6, 0, 'NULL')
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

    // 写文件
    const downWB = new Workbook();
    const downWS = downWB.addWorksheet('Sheet1');
    downWS.columns = columns
    downWS.addRows(allData);
    downWS.eachRow((row, rowIndex) => {
      row.eachCell(cell => {
        if (rowIndex === 1) {
          cell.style = {
            font: {
              size: 11,
              bold: true,
              color: {argb: 'ffffff'}
            },
            alignment: {vertical: 'middle', horizontal: 'center'},
            fill: {
              type: 'pattern',
              pattern: 'solid'
            },
            border: {
              top: {style: 'dashed', color: {argb: '0000ff'}},
              left: {style: 'dashed', color: {argb: '0000ff'}},
              bottom: {style: 'dashed', color: {argb: '0000ff'}},
              right: {style: 'dashed', color: {argb: '0000ff'}}
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
              top: {style: 'dashed', color: {argb: '0000ff'}},
              left: {style: 'dashed', color: {argb: '0000ff'}},
              bottom: {style: 'dashed', color: {argb: '0000ff'}},
              right: {style: 'dashed', color: {argb: '0000ff'}}
            }
          }
        }
      })
    })

    const fileName = 'temp_' + date + '_lx.xlsx'
    const filePath = uploadFolder + '/' + fileName;
    await downWB.xlsx.writeFile(filePath);


    const list = [{name: fileName, path: filePath}];
    res.render('filelist.ejs', {list: list, title: '佣金计算'})
  });

// 下载文件
router.get('/download2', function (req, res) {
  var filePath = req.query.path;
  console.log('下载文件：' + filePath)
  res.attachment(filePath)
  res.sendFile(filePath)
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

// 设置excel样式
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
const columns = [
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
// 3.4 勤勉 3单 630元.xlsx
const dpColumns = [
  {header: '日期', key: 'rq', width: 10},
  {header: '店铺名', key: 'dpm', width: 30},
  {header: '会员名', key: 'hym', width: 25},
  {header: '订单号', key: 'ddh', width: 25},
  {header: '价格', key: 'jg', width: 10},
  {header: '佣金', key: 'yj', width: 10}
]
// 【导出】店铺资金统计表.xlsx
const dpzjtjColumns = [
  {header: '日期', key: 'rq', width: 10},
  {header: '店铺名', key: 'dpm', width: 30},
  {header: '今日小计', key: 'jrxj', width: 20},
  {header: '昨日差额', key: 'zrce', width: 20},
  {header: '总计', key: 'zj', width: 20},
  {header: '公式', key: 'gs', width: 50},
  {header: 'Excel路径', key: 'lj', width: 10}
]


// 上传文件
router.post('/parse', upload.fields([{name: 'bdFile', maxCount: 1}]),
  async function (req, res) {
    if (!req.files || Object.keys(req.files).length === 0) {
      res.status(400).send('请选择要上传的文件！');
      return;
    }

    // 主持分组
    let cache = new Map();

    // budan
    const file = req.files.bdFile[0].path;
    const workbook = new Workbook();
    const workbook2 = await workbook.xlsx.readFile(file);

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
            jg: rowData[5].toString(), yj: rowData[6].toString(), zcyj: rowData[7].toString(), zc: rowData[8]
          }
          const dpObj = {
            rq: rowData[1], dpm: rowData[2], hym: rowData[3], ddh: rowData[4],
            jg: rowData[5].toString(), yj: rowData[6].toString()
          }
          var mapValue = cache.get(key);
          if (mapValue) {
            mapValue.push(bdObj)
          } else {
            cache.set(key, [bdObj])
          }
        }
      })
    })

    // 获取日期
    const originalname = req.files.bdFile[0].originalname
    const date = originalname.substring(5, originalname.indexOf('_', 5));
    // 前缀: 【导出】3.4
    const hhh = '【导出】' + date;
    // 遍历主持分组，写文件
    // 用来渲染页面的文件数组
    const list = [];
    for (let item of cache) {
      // 主持人：吴鑫
      const zc = item[0]
      const value = item[1]

      const downWB = new Workbook();
      const downWS = downWB.addWorksheet('Sheet1');
      downWS.columns = columns
      downWS.addRows(value);
      formatExcel(downWS);

      const dirPath = uploadFolder + '/' + zc;
      fs.rmdirSync(dirPath, {recursive: true})
      fs.mkdirSync(dirPath)

      // 文件名【导出】3.4博.xlsx
      const tempFileName = hhh + zc + '.xlsx'
      const filePath = dirPath + '/' + tempFileName;
      console.log(filePath);
      await downWB.xlsx.writeFile(filePath);

    //  主持下的店铺明细
      const dateDir = dirPath + '/' + hhh;
      fs.mkdirSync(dateDir)
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
        const thisName = date + ' ' + dpName + ' ' + tempList.length+'单 '+ totalPrice + '元.xlsx'
        tempList.push({})
        tempList.push({jg: '合计', yj: totalPrice})

        const thisWB = new Workbook();
        const thisWS = thisWB.addWorksheet('Sheet1');
        thisWS.columns = dpColumns
        thisWS.addRows(tempList);
        formatExcel(thisWS);
        const thisPath = dateDir + '/' + thisName;
        console.log(thisPath);
        await thisWB.xlsx.writeFile(thisPath);

        let zrceTemp = new Decimal(0)
        let zjTemp = sum.plus(zrceTemp).toFixed(2)
        let index = dpzjtjList.length + 2
        dpzjtjList.push({
          rq: date, dpm: dpName, jrxj: totalPrice, zrce: '0', lj: '',
          zj: {formula: `=C${index}+D${index}`},
          gs: {formula: `=A${index}&"  今日合计"&C${index}&"， 昨日差额"&D${index}&"，  合计差额"&E${index}`}
        })
      }

      const dpzjWB = new Workbook();
      const dpzjWS = dpzjWB.addWorksheet('Sheet1');
      dpzjWS.columns = dpzjtjColumns
      dpzjWS.addRows(dpzjtjList);
      dpzjtjFormatExcel(dpzjWS);
      const dpzjPath = dirPath + '/【导出】店铺资金统计表.xlsx';
      console.log(dpzjPath);
      await dpzjWB.xlsx.writeFile(dpzjPath);


      // 创建 zip
      const zipFile = new AdmZip();
      zipFile.addLocalFolder(dirPath, zc);
      const zipName = '_' + zc + '.zip'
      const zipPath = uploadFolder + '/' + zipName
      zipFile.writeZip(zipPath);
      list.push({name: zipName, path: zipPath})
    }


    res.render('filelist.ejs', {list: list, title: '表格拆分'})
  });


// 上传页面
router.get('/index', (req, res) => {
  console.log(__dirname)
  res.sendFile(path.join(__dirname, '../views/upload.html'))
})

// 列表页面
router.get('/filelist', function (req, res) {
  res.sendFile(path.join(__dirname, '../views/filelist.html'))
})

// 上传文件
router.post('/upload', upload.array('file'), function (req, res) {
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
