var xlsx = require("node-xlsx");
var fs = require('fs');
var list = xlsx.parse("test.xlsx");
praseExcel(list);
//解析Excel
function praseExcel(list)
{
    console.log("start...");
    for (var i = 0; i < list.length; i++) 
    {
         var excleData = list[i].data;
         var sheetArray  = [],keyarr=[],zharr=[],enarr=[];
         var typeArray =  excleData[1];
         // var keyArray =  excleData[2];
         var allObj = {
         	zh:{},
         	en:{}
         };
        for (var j = 2; j < excleData.length ; j++)
        {
             var curData = excleData[j];
             if(curData.length == 0) continue;
             var mykey = changeObj(curData,typeArray,'key');
             var myzh = changeObj(curData,typeArray,'zh');
             var myen = changeObj(curData,typeArray,'en');
             keyarr.push(mykey);
             zharr.push(myzh);
             enarr.push(myen);
             }
        for(var k = 0; k < zharr.length; k++){
        	var item = {};
        	item[keyarr[k]] = zharr[k];
        	// console.log(item)
        	sheetArray.push(item);
        	allObj.zh[keyarr[k]] = zharr[k];
        	allObj.en[keyarr[k]] = enarr[k];
        }
        if(sheetArray.length >0) 
        writeFile(list[i].name+".json",JSON.stringify(allObj));
    }
}
//转换数据类型
function changeObj(curData,typeArray,flag)
{
     var obj = '',arr1=[],arr2=[];
     // console.log(typeArray)
    for (var i = 0; i < curData.length; i++) 
    {
        var myvalue = changeValue(curData[i],typeArray[i],flag);
        obj = myvalue;
        if(myvalue != ''){
        	return myvalue;
        }
    }
    return obj;
}
function changeValue(value,type,flag)
{
    if(value == null || value =="null"||type=='') return ""
	if(type == flag){
		return value;
	}else{
		return "";
	}
}
//写文件
function writeFile(fileName,data)
{  
  fs.writeFile(fileName,data,'utf-8',complete);
  function complete(err)
  {
      if(!err)
      {
          console.log("文件生成成功");
      }   
  } 
}