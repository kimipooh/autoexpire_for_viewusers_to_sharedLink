var AfterDay = 7; // 申請日から○日後のパラメーター
var URL = "https://drive.google.com/open?id=ファイル（フォルダ）ID";

// AfterDay後にトリガーとして読み出し（共有の解除）
function autoExpireSharedUsers(){
  var id, asset;

  try{
    var id = URL.split('=')[1];

    if(id){
      // ファイル、フォルダいずれでもOKにする
      asset = DriveApp.getFileById(id) ? DriveApp.getFileById(id) : DriveApp.getFolderById(id);
      
      if(asset){
        // シート全体から、「現在の年月日 = (A列のタイムスタンプ + AfterDay後)の年月日」がマッチする行のデータを抽出
        var unlinked_rows = getTimestampFromSpreadsheet();
        if(unlinked_rows){
          for (var i in unlinked_rows) {
            // マッチした行のメールアドレスを取得
            var email = unlinked_rows[i][1];
            if(email){           
              // 読み取りユーザーを削除
              asset.removeViewer(email);
            }
          }
        }      
      }
    }
    
  } catch (e) {
    Logger.log(e.toString());    
  }

}

// シート全体から、「現在の年月日 = (A列のタイムスタンプ + AfterDay後)の年月日」がマッチする行のデータを抽出する関数
function getTimestampFromSpreadsheet() {
//有効なGooglesプレッドシートを開く
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//全セルを保存
 var data = sheet.getDataRange().getValues();
//抽出データ
 var unlinked_rows = [];

 if(data){
   for (var num_row in data) {
     //タイムスタンプ（申請日時）を取得
     var timestamp = new Date(data[num_row][0]);

     if(!isNaN(timestamp.getTime())){
       //メールアドレスの取得
       var email = data[num_row][1];
       // タイムスタンプから、７日後の年月日を計算する
       var timestamp_afterday = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate() + AfterDay);
       var ta_date = timestamp_afterday.getFullYear() + "" + timestamp_afterday.getMonth() + "" + timestamp_afterday.getDate(); 

       // 今の年月日
       var current_date = new Date();
       var cu_date = current_date.getFullYear() + "" + current_date.getMonth() + "" + current_date.getDate(); 

       // 申請時のタイムスタンプ + AfterDay 日後が、本日であるなら、日時を保存
       if(ta_date === cu_date){
         unlinked_rows.push(data[num_row]); 
       }
       
     }
   }                       
 }

 return unlinked_rows;
}

// フォーム送信時に呼び出し
function autoexpire_for_viewusers_to_sharedLink() {
//有効なGooglesプレッドシートを開く
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//新規申請された行番号を取得
 var num_row = sheet.getLastRow();
//タイムスタンプ（申請日時）を取得
 var timestamp = new Date(sheet.getRange(num_row, 1).getValue());
//メールアドレスの取得
 var email = sheet.getRange(num_row,2).getValue();
//タイムスタンプを年月日に分解する
 // 月、日が一桁なら 0 を先頭に追加する関数の定義（ 9 --> 09 等）
 var toDoubleDigits = function(num) {
  num += "";
  if (num.length === 1) {
    num = "0" + num;
  }
  return num;     
 };
// タイムスタンプから、AfterDay 日後の年月日を計算する
 var timestamp_afterday = new Date(timestamp.getFullYear(), timestamp.getMonth(), timestamp.getDate() + AfterDay); 
// メール通知用： タイムスタンプから、７日後の年月日を取り出す（年 = yyyy, 月 = mm, 日 = dd に保存する）
 var yyyy = timestamp.getFullYear();
 var mm   = toDoubleDigits(timestamp.getMonth()+1);
 var dd   = toDoubleDigits(timestamp.getDate());

 var id, asset;

 try{
   var id = URL.split('=')[1];

   if(id && email){
     asset = DriveApp.getFileById(id) ? DriveApp.getFileById(id) : DriveApp.getFolderById(id);
     
     // https://developers.google.com/apps-script/reference/drive/file#setSharing(Access,Permission)
      if(asset){
        // https://developers.google.com/apps-script/reference/drive/access.html
        // https://developers.google.com/apps-script/reference/drive/permission.html
        // https://developers.google.com/apps-script/reference/drive/file.html#addViewer(String)
        asset.addViewer(email);
      }   
     // メール通知設定
     var contents = "ダウンロード先と有効期限は次の通りです。 \n\n"
     +"有効期限： "+yyyy+"年"+mm+"月"+dd+"日"
     +"\n\nダウンロードリンク: "+asset.getUrl()
     +"\n\nGoogleアカウント（" + email + "） でアクセスしてください。";

     // タイムスタンプから、AfterDay 日後、autoExpireSharedUsers（共有解除）関数を実行するようトリガーに追加
     if ( !isNaN (timestamp_afterday.getTime()) ){
       ScriptApp.newTrigger("autoExpireSharedUsers").timeBased().at(timestamp_afterday).create();
     }
     // メール送信
     MailApp.sendEmail(email,"ダウンロードリンクのお知らせ",contents);

   }
 } catch (e) {
    Logger.log(e.toString());    
 }
}