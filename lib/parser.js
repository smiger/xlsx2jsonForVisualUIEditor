const config = require('../config.json');
const types = require('./types');
var _ = require('lodash');

const DataType = types.DataType;
const SheetType = types.SheetType;

var CID = 1;
var UICheckBoxCount = 0;
var CKArr = [];
var TIPCount = 0;
var TIPArr = [];
var ITEM_Y = 0;

function parseJsonObject(data) {
  var evil = eval;
  return evil("(" + data + ")");
}

// class StringBuffer {
//   constructor(str) {
//     this._str_ = [];
//     if (str) {
//       this.append(str);
//     }
//   }

//   toString() {
//     return this._str_.join("");
//   }

//   append(str) {
//     this._str_.push(str);
//   }
// }


/**
 * 解析workbook中所有sheet的设置
 * @param {*} workbook 
 */
function parseSettings(workbook) {
  let settings = {};

  workbook.forEach(sheet => {
    //叹号开头的sheet不输出
    if (sheet.name.startsWith('!')) {
      return;
    }

    let sheet_name = sheet.name;
    let head_row = sheet.data[0];

    let sheet_setting = {
      type: SheetType.NORMAL,
      master: null,
      slaves: [],
      ids:[],
      head: []
    };
    if (sheet_name === config.xlsx.master) {
      sheet_setting.type = SheetType.MASTER;
    }else{
      sheet_setting.type = SheetType.SLAVE;
    }
    //parsing head setting
    head_row.forEach((cell,index) => {

      cell = cell.toString();
      let t = sheet.data[config.xlsx.head+1][index];
      let head_setting = {
        name: cell,
        type: t,
      };
      
      if (cell.indexOf(':#') !== -1) {
        let pair = cell.split(':#');
        let name = pair[0].trim();
        let id = pair[1].split('.')[1].trim();
        head_setting.name = name;

        sheet_setting.master = sheet_name;
        sheet_setting.slaves.push(name);
        sheet_setting.ids.push(id);
      }
      
      sheet_setting.head.push(head_setting);
    });

    settings[sheet_name] = sheet_setting;
  });
  return settings;
}

/**
 * 解析一个表(sheet)
 *
 * @param sheet 表的原始数据
 * @param setting 表的设置
 * @return Array or Object
 */
function parseSheet(sheet, setting) {

  let headIndex = config.xlsx.head + 2;
  let result = [];

  if (setting.type === SheetType.MASTER) {
    result = {};
  }

  console.log('  * sheet:', sheet.name, 'rows:', sheet.data.length);

  UICheckBoxCount = 0;
  CKArr = [];
  TIPArr = [];
  CID = 1;
  TIPCount = 0;
  for (let i_row = headIndex; i_row < sheet.data.length; i_row++) {

    let row = sheet.data[i_row];

    let parsed_row = parseRow(row, i_row, setting.head);

    if (setting.type === SheetType.MASTER) {

      let id_cell = setting.head[0];
      if (!id_cell) {
        throw `在表${sheet.name}中获取不到id列`;
      }

      result[parsed_row[id_cell.name]] = parsed_row;
      // result = parsed_row;

    } else {
      result.push(parsed_row);
    }
  }
  CKArr.push(UICheckBoxCount);
  TIPArr.push(TIPCount);
  console.log("UICheckBoxCount:" + TIPArr);
  CID = 1;
  if(CKArr.length > 0){
    let n = Math.floor((CKArr[0]-1)/3);
    let inc = 0;
    for(let i = 0; i < result.length; i++){
      let row = result[i];
      if(row['cId'] != CID){
        inc = 0;
        CID = row['cId'];
        CKArr.shift();
        n = Math.floor((CKArr[0]-1)/3);
        TIPArr.shift();
      }
      //item项
      if(row['type'] && row['type'] == 'UICheckBox'){
        let m = Math.floor(inc/3);
        if(n-m >= 0){
          row['y'] = row['y'] + 55*(n-m);
        }
        inc++;
        row['touchEnabled'] = true;
        row['width'] = 44 + "";
        row['height'] = 44 + "";
        //按钮类型
        if(row['btnType']){
          if(row['btnType']=='CheckBox'){
            row['back']='uires/main/btn_checkbox_fang.png';
            row['active']='uires/main/btn_checkbox_fang1.png';
            row['select'] = true;
          }else if(row['btnType'] == 'RadioButton'){
            row['back']='uires/main/btn_checkbox.png';
            row['active']='uires/main/btn_checkbox1.png';
            row['select'] = true;
          }
          delete row['btnType'];
        }
        //有描述
        if(row['desc']){
          let desc = {};
          desc['cId'] = row['cId'];
          let str = row['id'] + '';
          var last = str.match(/(\d+)$/g)[0]
          // let last = str.substring(str.length - 1);
          str = str.replace('ck_','lab_');
          str = str.replace(last,'_'+last);
          desc['id'] = str;
          desc['touchEnabled'] = false;
          desc['x'] = row['x'] + 30;
          desc['y'] = row['y'];
          if(TIPArr[0]){
            desc['y'] = desc['y'] + 25;
          }
          if(row['detail']){
            desc['y'] = desc['y'] + 10;
          }
          desc['type'] = "UIText";
          desc['anchorX'] = 0;
          desc['string'] = row['desc'];
          desc['fontSize'] = 24;
          desc['fontName'] = "Arial";
          result.splice(i+1,0,desc);
          delete row['desc'];
        }
        //有提示
        if(row['tip']){
          let tip = {};
          tip['cId'] = row['cId'];
          tip['id'] = row['tipid'];
          tip['touchEnabled'] = false;
          tip['x'] = 456;
          tip['y'] = 17;
          tip['type'] = "UIText";
          tip['color'] = [
              0,
              255,
              255,
              255
          ];
          tip['string'] = row['tip'];
          tip['fontSize'] = 16;
          tip['fontName'] = "Arial";
          result.push(tip);
          delete row['tip'];
          delete row['tipid'];
          row['y'] = row['y'] + 25;
        }
        //小提示
        if(row['detail']){
          let tip = {};
          tip['cId'] = row['cId'];
          tip['id'] = row['tipid'];
          tip['touchEnabled'] = false;
          tip['x'] = row['x']+30;
          tip['y'] = row['y']-14;
          tip['type'] = "UIText";
          tip['anchorX'] = 0;
          tip['string'] = row['detail'];
          tip['fontSize'] = 12;
          tip['fontName'] = "Arial";
          result.push(tip);
          delete row['detail'];
          delete row['tipid'];
        }
      }
      //item背景UIScale9
      if(row['type'] && row['type'] == 'UIScale9'){
        row['height'] = 60 + 55 * n + "";
        if(TIPArr[0]){
          row['height'] = parseInt(row['height']) + 25 + '';
        }
        row['spriteFrame']="uires/main/bg_neirongdiwen.png";
        row["insetLeft"]= 10;
        row["insetTop"]= 10;
        row["insetRight"]= 10;
        row["insetBottom"]= 10;
      }
    }
    
  }
  return result;
}
/**
 * 解析一行
 * @param {*} row 
 * @param {*} rowIndex 
 * @param {*} head 
 */
function parseRow(row, rowIndex, head) {

  let result = {};
  let id;

  for (let index = 0; index < head.length; index++) {
    let cell = row[index];

    let name = head[index].name;
    let type = head[index].type;

    if (name.startsWith('!')) {
      continue;
    }

    if (cell === null || cell === undefined) {
      // result[name] = null;
      continue;
    }

    if(cell == 'UICheckBox'){
      if(result['cId'] != CID){
        CKArr.push(UICheckBoxCount);
        UICheckBoxCount = 0;
        CID = result['cId'];
        TIPArr.push(TIPCount);
        TIPCount=0;
      }
      if(!result['x']){
        result['x'] = 150 + 250 * Math.ceil(UICheckBoxCount%3);
        result['y'] = 30;
        UICheckBoxCount++;
      }
      if(result['tip']){
        TIPCount++;
      }
    }else if(cell == 'UIScale9'){
      result['x'] = 103;
      result['y'] = 0;
      result['width'] = 720+"";
      result['height'] = 60+"";
    }
    switch (type) {
      case DataType.ID:
        id = cell + '';
        result[name] = id;
        break;
      case DataType.IDS:
        id = cell + '';
        result[name] = id;
        break;
      case DataType.UNKNOWN: // number string boolean
        if (isNumber(cell)) {
          result[name] = Number(cell);
        } else if (isBoolean(cell)) {
          result[name] = toBoolean(cell);
        } else {
          result[name] = cell;
        }
        break;
      case DataType.DATE:
        if (isNumber(cell)) {
          //xlsx's bug!!!
          result[name] = numdate(cell);
        } else {
          result[name] = cell.toString();
        }
        break;
      case DataType.STRING:
        if (cell.toString().startsWith('"')) {
          result[name] = parseJsonObject(cell);
        } else {
          result[name] = cell.toString();
        }
        break;
      case DataType.NUMBER:
        //+xxx.toString() '+' means convert it to number
        if (isNumber(cell)) {
          result[name] = Number(cell);
        } else {
          console.warn("type error at [" + rowIndex + "," + index + "]," + cell + " is not a number");
        }
        break;
      case DataType.BOOL:
        result[name] = toBoolean(cell);
        break;
      case DataType.OBJECT:
        result[name] = parseJsonObject(cell);
        break;
      case DataType.ARRAY:
      case DataType.ARRAY2:
        if (!cell.toString().startsWith('[')) {
          cell = `[${cell}]`;
        }
        result[name] = parseJsonObject(cell);
        break;
      default:
        console.log('无法识别的类型:', '[' + rowIndex + ',' + index + ']', cell, typeof (cell));
        break;
    }
  }

  return result;
}

/**
 * convert value to boolean.
 */
function toBoolean(value) {
  return value.toString().toLowerCase() === 'true';
}

/**
 * is a number.
 */
function isNumber(value) {

  if (typeof value === 'number') {
    return true;
  }

  if (value) {
    return !isNaN(+value.toString());
  }

  return false;
}

/**
 * boolean type check.
 */
function isBoolean(value) {

  if (typeof (value) === "undefined") {
    return false;
  }

  if (typeof value === 'boolean') {
    return true;
  }

  let b = value.toString().trim().toLowerCase();

  return b === 'true' || b === 'false';
}

//fuck node-xlsx's bug
var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
// var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
// function datenum(v, date1904) {
// 	var epoch = v.getTime();
// 	if(date1904) epoch -= 1462*24*60*60*1000;
// 	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
// }

function numdate(v) {
  var out = new Date();
  out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
  return out;
}
//fuck over

function parseSalves(name,parsed_workbook,settings){
  let master_sheet = parsed_workbook[name];
  settings[name].slaves.forEach((slave_name,index) => {

    let slave_setting = settings[slave_name];
    let slave_sheet = parsed_workbook[slave_name];
    let mid = settings[name].ids[index];
    let key_cell = _.find(slave_setting.head, item => {
      return item.name === mid;
    });
    let key_slave = _.find(settings[name].head, item => {
      return item.name === slave_name;
    });
    console.log('-------------------------'+slave_name + 'index:'+index);
    //slave 表中所有数据
    slave_sheet.forEach(row => {
      let id = row[key_cell.name];
      delete row[key_cell.name];
      if(!id){
        return false;
      }
      let key_master_sheet = _.find(master_sheet, item => {
        return item[mid] === id;
      });
      //判断是否grp子项
      if(key_master_sheet['type']&&key_master_sheet['type']=='UIWidget'&&key_master_sheet['id']!='grp_item'){
        if(row['type'] == 'UIScale9'){
          key_master_sheet['height']=row['height'];
          key_master_sheet['width']=row['width'];
          key_master_sheet['x']=0;
          key_master_sheet['y']= ITEM_Y;
          ITEM_Y = ITEM_Y + parseInt(row['height']) + 6;
          if(key_master_sheet['desc']){
            let desc = {};
            desc['cId'] = row['cId'];
            desc['touchEnabled'] = false;
            desc['x'] = 65;
            desc['y'] = parseInt(row['height'])/2+ "";
            desc['type'] = "UIText";
            desc['string'] = key_master_sheet['desc'];
            desc['fontSize'] = 24;
            desc['fontName'] = "Arial";
            key_master_sheet["children"] = key_master_sheet["children"] || [];
            key_master_sheet["children"].push(desc);
            delete key_master_sheet['desc'];
          }
        }
      }
      
      if (key_cell.type === DataType.IDS || key_slave.type === DataType.ARRAY2) { //array
        key_master_sheet["children"] = key_master_sheet["children"] || [];
        key_master_sheet["children"].push(row);
      } else { //hash
        key_master_sheet["children"] = row;
      }
    });
    if(settings[slave_name].master){
      parseSalves(settings[slave_name].master,parsed_workbook,settings);
      delete parsed_workbook[slave_name];
    }else{
      delete parsed_workbook[slave_name];
    }
  });
  return parsed_workbook;
}

module.exports = {

  parseSettings: parseSettings,
  parseWorkbook: function (workbook, settings) {

    // console.log('settings >>>>>', JSON.stringify(settings, null, 2));

    let parsed_workbook = {};

    workbook.forEach(sheet => {

      if (sheet.name.startsWith('!')) {
        return;
      }

      let sheet_name = sheet.name;

      let sheet_setting = settings[sheet_name];

      let parsed_sheet = parseSheet(sheet, sheet_setting);

      parsed_workbook[sheet_name] = parsed_sheet;

    });
    console.log("-----------------------")
    for (let name in settings) {
      if (settings[name].type === SheetType.MASTER) {
        parsed_workbook = parseSalves(name,parsed_workbook,settings);
      }
    }

    return parsed_workbook;
  }
};