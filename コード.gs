function GmailDistribution() {
  //EST();

  //署名の作成
  let signature_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('署名');//署名シート
  let signature = CreateSentences(signature_sheet, 1); //署名シート１行目～最終行目を改行で結合した文字列
  
  //各シートに作成された本文を取得し、message_objに格納
  let last_sheet = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();//4
  let message_obj={};　//空のobjを設置
  /*
    message_obj = { 本文設定用のシート名：{’subject’:件名（1行目）,'message':本文（2～最終行目）} , ... }
   */
  for( let i=0; i<last_sheet; i++ ){ //1シート目～最終シートまでループ
    let sheet_name = SpreadsheetApp.getActiveSpreadsheet().getSheets()[i].getSheetName();//対象シート名を取得
    if(sheet_name=='送信先' || sheet_name=='署名'){//本文を設定したシートではないので処理なし
      continue; //何もせず次のループへ
    }else{　//本文設定用のシート
      //message_objに情報を格納
      GetMessage(sheet_name, message_obj);
    }
  }
  //送信先シートを走査し、対象者へメール作成＆送信
  let address_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('送信先');//送信先シート
  let last_address = address_sheet.getLastRow();//送信先シートの最終行
  //送信先シートの2~最終行目までを１行ずつループ
  for( let i=2; i<=last_address; i++ ){
      let name = address_sheet.getRange(i,1).getValue(); //A列：氏名
      let email_address = address_sheet.getRange(i,2).getValue(); //B列：メールアドレス
      let message_category = address_sheet.getRange(i,3).getValue(); //C列：本文　…　シート名に対応
      let subject = message_obj[message_category]['subject'] //件名
      let body = name +' 様'+'\n'+message_obj[message_category]['message'] + '\n\n' + signature //宛名 +  様 + 改行 + 本文 + 改行*2 + 署名

      //GmailApp.sendEmail(email_address, subject, body);
    //TEST用
      GmailApp.createDraft(email_address, subject, body);

    }
}

function GetMessage(sheet_name, message_obj){
  let target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);

  let inner_obj={};
  inner_obj['subject'] = target_sheet.getRange(1,2).getValue();//件名を格納
  //本文
  inner_obj['message'] = CreateSentences(target_sheet, 2);//本文を格納

  //シート名をキーに、message_objに格納
  message_obj[sheet_name] = inner_obj;

}

function CreateSentences(target_sheet, first){
  let last = target_sheet.getLastRow();//対象シート最終行番号
  sentences="";
  for(let i=first; i<=last; i++){　//対象シートのfirst行から最終行まで１行する走査
    //文字列sentencesにB列の値+改行コードを付け足していく
    sentences += target_sheet.getRange(i,2).getValue() + '\n';
  }
  return sentences
}

function TEST(){
  for(let i=0; i<4; i++){
    let result=SpreadsheetApp.getActiveSpreadsheet().getSheets()[i].getSheetName();
    console.log(result);
  }


}


  
