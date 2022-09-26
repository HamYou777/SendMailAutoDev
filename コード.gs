function getDate() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("シート1");

  const firstRow = 6;  
  const dayCol = 1; 
  const corpCol = 2;
  const groupCol = 3;
  const nameCol = 4;
  const mailCol = 5; 
  const flagCol = 6;
  
  let today = new Date();
  today = Utilities.formatDate(today, "JST", "yyyy/MM/dd");
  let rowLength = sheet.getLastRow() - (firstRow - 1);
  let dayList = [];
  let corpList = [];
  let groupList = [];
  let nameList = [];
  let mailaddressList = [];
  let flagList = [];


  for(let i=0; i < rowLength; i++) {
    // 日付取得
    let day = sheet.getRange(firstRow + i, dayCol).getValue();
    if(day == "") {
      continue;
    }    
    formattedDay = Utilities.formatDate(day, "JST", "yyyy/MM/dd");
    if(formattedDay != today) {
      continue;
    }
    dayList.push(formattedDay);
    // 会社名取得
    let corp = sheet.getRange(firstRow + i, corpCol).getValue();
    corpList.push(corp);
    // 所属取得
    let group = sheet.getRange(firstRow + i, groupCol).getValue();
    groupList.push(group);
    // 名前取得
    let name = sheet.getRange(firstRow + i, nameCol).getValue();
    nameList.push(name);
    // メールアドレスリスト取得
    let address = sheet.getRange(firstRow + i, mailCol).getValue();
    mailaddressList.push(address);
    // 結果をオブジェクトとして取得
    let flag = sheet.getRange(firstRow + i, flagCol);
    flagList.push(flag);
  }


  mailaddressList.forEach((address, index) => {
    var message = "";
    if(corpList[index] == "" || nameList[index] == "" || address == "") {
      var message = "入力に不備があり、メールは送信されませんでした。";
      console.log (
        "入力に不備があり、メールは送信されませんでした。address => " +  address +
        ", 日付 => " + dayList[index] + ",　会社名　=> " + corpList[index] + ",　所属　=> " +   groupList[index] + ",　名前　=> " + nameList[index]
      );
    }
    else {
      let corp = corpList[index];
      let group = groupList[index];
      let name = nameList[index];
      sendMail(address,corp,group,name);
      var message = ("送信済み");
      console.log(
        "メールが送信されました。address => " + address + 
        ", 日付 => " + dayList[index] + ",　会社名　=> " + corpList[index] + ",　所属　=> " +   groupList[index] + ",　名前　=> " + nameList[index]
      );
    }
    // 結果へ出力
      flagList[index].setValue(message);
  });
}

function sendMail(address, corp, group, name) {
  const subject = "【全体連絡】システム休止日について"; // メールの件名

  const bodyTemplate = `
{corp} {group} {name}様

総務部より社内システム休止日についてお知らせです。
社内システムの入れ替えに伴い、20XX年〇月〇日は終日社内システムが休止となります。
関係部署においては業務調整をよろしくお願いいたします。

詳細内容については社内イントラに掲示していますので、ご一読ください。
https://〇〇〇〇〇〇〇〇〇〇〇/

各種問い合わせは部門内で取りまとめの上、総務部までお願いいたします。

総務部　社内システム担当　〇〇
内線：000-000-0000

`;

  let body = bodyTemplate
        .replace("{corp}", corp).replace("{group}", group).replace("{name}", name);

  // GmailApp.sendEmail(address, subject, body);

}