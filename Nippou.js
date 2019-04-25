// N列に「日報報告トリガー」、O列に「報告済み」と記入をします。
//トリガーは12~1時を指定する。
//以下はメンバー各々が設定します。
var API_TOKEN = 'API';　//自身のチャットワークAPIトークン
var name = "John";  //本名フルネームで
var roomId = "XXX"; //プログラミングコースのルームID
var myroomId = "YYY"; //マイチャットのルームID
var myids = "BBB"; //自分のChatworkID。自分宛のタスクのaid＝"xxx"や自分宛のto="xxx"をメッセージの編集画面で確認してください。
var mgids = "CCC";  //マネージャーのChatworkID。マネージャー宛のタスク：aid＝"xxx"や自分宛のto="xxx"をメッセージの編集画面で確認してください。
var ids ="BBB,CCC"; //上二つのIDをカンマで区切って記入してください。
var spreadsheet = SpreadsheetApp.openById('ID');　　// SPreadsheetのURLのd/XXX/editにおけるXXXの部分
var sheetName = 'John'; //達成シート状の自分のシートの名前名をコピペしてください。

var params1 = {
    headers : {"X-ChatWorkToken" : API_TOKEN},
    method : "post"
  };

var params2 = {
    headers : {"X-ChatWorkToken" : API_TOKEN},
    method : "post"
　}

function doPost() {
  
  var sheet = spreadsheet.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  var target = 1;
  var targetColumnIndex = 13;
  var defineColumnIndex = 14;
  var targetRowIndexes = [];
  
  //Chatwork-Taskに時間指定をする
  var dates = new Date();
  var limit = dates.getTime()/1000; 
  limit = limit.toFixed();
  
  var searchDate = Utilities.formatDate(new Date(), 'Asia/Tokyo','yyyy/MM/dd');
  
  for (var i = data.length - 1; i > 0; i--){
    var dataDate = Utilities.formatDate(new Date(data[i][1]), 'Asia/Tokyo','yyyy/MM/dd');
    if((dataDate == searchDate) && (data[i-1][3] != "")){
      targetRowIndexes.push(i-1);
      break;    
    }
  }
 
  var targetRow = targetRowIndexes[0];
  
  //データがなかったら終了
  if(!targetRow){
    return;
  }
  
  var body = '';
  var errorbody = '';
  var theday = data[targetRow][1];
  var month = theday.getMonth() + 1;
  var date = theday.getDate() ;
  var expectedAim = data[targetRow][3];
  var aim = data[targetRow][4];
  var taisaku = data[targetRow][6];
  var yotei = data[targetRow][7];
  var zissai　= data[targetRow][8];
  var gakusyu = data[targetRow][9];
  
  try{
      
    if((data[targetRow][targetColumnIndex] == "完了です！")　&& (!data[targetRow][defineColumnIndex])){
    
    body += "名前：" + name;
    body += "\n日付：" +month + "/" + date;
    body += "\n目標：" + aim;   
    
    if(aim == "未達"){
      body += "\n対策：" + taisaku;
      if(taisaku == ""){
        errorbody +=　"＊" + month + "/" + date +　"の補填日を記入してください。"
      }
    }else if(aim == ""){
      if(errorbody == ""){
       errorbody +=　"＊"+ month + "/" + date + "の達成・未達成を記入してください。"
      }else{
       errorbody +=　"\n＊"+ month + "/" + date + "の達成・未達成を記入してください。"
      }
    }
     
    body += "\n予定学習時間：" + yotei;
    body += "\n実際学習時間：" + zissai;
    
    if(yotei == ""){
      if(errorbody == ""){
        errorbody +=　"＊" + month + "/" + date +　"の予定学習時間を記入してください。"
      }else{
        errorbody +=　"\n＊" + month + "/" + date +　"の予定学習時間を記入してください。"
      }
    }
     
    if(zissai == ""){
      if(errorbody == ""){
        errorbody +=　"＊" + month + "/" + date +　"の実際学習時間を記入してください。"
      }else{
        errorbody +=　"\n＊" + month + "/" + date +　"の実際学習時間を記入してください。"
      }
    }
   
    if(yotei > zissai){
      if(errorbody == ""){
      errorbody += "＊" + month + "/" + date + "の実際学習時間が予定学習時間より少ないです。達成シートはあっていますか？"
      }else{
      errorbody += "\n＊" + month + "/" + date + "の実際学習時間が予定学習時間より少ないです。達成シートはあっていますか？"
      }
    }
    
    if(zissai > 0 && (gakusyu == "")){
      if(errorbody == ""){
      errorbody += "＊勉強されたようですね！学習内容も記入しましょう！"
      }else{
      errorbody += "\n＊勉強されたようですね！学習内容も記入しましょう！"
      } 
    }
    
    if(errorbody.length >= 1){   
      var url2 = "https://api.chatwork.com/v2/rooms/" + roomId + "/tasks";
      params2.payload = {body : errorbody, limit : limit, to_ids : ids};
      UrlFetchApp.fetch(url2, params2);    
    }
  
  var url = "https://api.chatwork.com/v2/rooms/" + roomId + "/tasks";
  params1.payload = {body :body, limit : limit, to_ids : mgids};
  UrlFetchApp.fetch(url, params1);
  sheet.getRange(targetRow+1, defineColumnIndex+1,1,1).setValue('報告済みです！');
  }else{
      errorbody += "＊"  + month + "/" + date + "の達成シートが記入されていません。記入してください！"
      var url2 = "https://api.chatwork.com/v2/rooms/" + roomId + "/tasks";
      params2.payload = {body : errorbody, limit : limit, to_ids : ids};
      UrlFetchApp.fetch(url2, params2);  
  }
  }
    catch(e){
    var errorbody = "＊上手く処理できませんでした。担当者に報告してください。"
    var url2 = "https://api.chatwork.com/v2/rooms/" + myroomId + "/tasks";
    params2.payload = {body : errorbody, limit : limit, to_ids : myids};
    UrlFetchApp.fetch(url2, params2);   
  }
}
