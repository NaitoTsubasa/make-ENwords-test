// toppageurlはresult.htmlの"作成画面に戻る"で使用されます. グローバル変数でないといけません.
var toppageurl = ScriptApp.getService().getUrl();

function doGet(){
  var toppage = HtmlService.createTemplateFromFile("web").evaluate();
  toppage.setTitle("英単語テスト作成");
  toppage.setFaviconUrl("--------------------------------------")
  return toppage;
}

function doPost(e){
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  
  var folder = DriveApp.getFolderById("--------------------------------");

  if(!e.parameters.page){
    var error = HtmlService.createTemplateFromFile("error");
    error.toppageurl = toppageurl;
    error = error.evaluate().setTitle("作成エラー").setFaviconUrl("----------------------------------------");
    Logger.log("エラー処理実行");
    return error;
  }
  var docetoj = DocumentApp.create("英単語テスト(英→日)" + timestamp);
  var docjtoe = DocumentApp.create("英単語テスト(日→英)" + timestamp);

  switch(e.parameter.grade.toString()){
    case "フォレスタ1年生":
      var spreadsheet = SpreadsheetApp.openById("----------------------------");
      Logger.log("フォレスタ1年生");
      break;
    case "フォレスタ2年生":
      var spreadsheet = SpreadsheetApp.openById("------------------------------------");
      Logger.log("フォレスタ2年生");
      break;
    case "フォレスタ3年生":
      var spreadsheet = SpreadsheetApp.openById("-----------------------------------------");
      Logger.log("フォレスタ3年生");
      break;
  }

  docetoj.addHeader().setText(e.parameter.grade.toString() + '\t' + e.parameters.page.toString());
  docjtoe.addHeader().setText(e.parameter.grade.toString() + '\t' + e.parameters.page.toString());

  var pages = e.parameters.page.toString();
  pages = pages.split(",");
  Logger.log(pages);
  var en = [];
  var ja = [];
  var contents = "　　";

  for(let i=0;i<pages.length;i++){
    var sheet = spreadsheet.getSheetByName(pages[i]);
    var lastrow = sheet.getLastRow();
    var enlist = sheet.getRange(1,1,lastrow).getValues();
    en = en.concat(enlist);
    var jalist = sheet.getRange(1,2,lastrow).getValues();
    ja = ja.concat(jalist);
    var pagetitle = sheet.getRange("C1").getValue();
    contents += pagetitle + "　　";
  }

  en = Array.prototype.concat.apply([],en);
  ja = Array.prototype.concat.apply([],ja);

  var encontents = "";
  var jacontents = "";

  switch(e.parameter.shuffle.toString()){
    case "False":
      var shuffleqs = "シャッフルしない"
      Logger.log("シャッフルしない")
      break;
    case "True":
      var shuffleqs = "シャッフルする"
      en = shuffle(en);
      ja = shuffle(ja);
      Logger.log("シャッフルする")
      break;
    case "shuffle33":
      var shuffleqs = "シャッフル33問（1ページ）"
      en = shuffle(en)
      ja = shuffle(ja)
      if(en.length >= 33){
        en = en.slice(0,32);
        ja = ja.slice(0,32);
      }
      Logger.log("シャッフル33問")
      break;
    case "shuffle66":
      var shuffleqs = "シャッフル66問（2ページ）"
      en = shuffle(en)
      ja = shuffle(ja)
      if(en.length >= 66){
        en = en.slice(0,65);
        ja = ja.slice(0,65);
      }
      Logger.log("シャッフル66問")
      break;
  }

  var logsheets = SpreadsheetApp.openById("------------------------------------");
  var logsheet = logsheets.getSheets()[0];
  var loglastrow = logsheet.getLastRow();
  logsheet.getRange(loglastrow+1,1).setValue(timestamp);
  logsheet.getRange(loglastrow+1,2).setValue(shuffleqs);
  logsheet.getRange(loglastrow+1,3).setValue(e.parameter.grade.toString());
  logsheet.getRange(loglastrow+1,4).setValue(e.parameters.page.toString());

  for(let i=0;i<en.length;i++){
    encontents += en[i];
    for(let j=29;j>en[i].length;j--){
      encontents += " ";
    }
    encontents += "_________________\n"
  }

  for(let i=0;i<ja.length;i++){
    jacontents += ja[i];
    for(let j=21;j>ja[i].length;j--){
      jacontents += "　";
    }
    jacontents += "＿＿＿＿＿＿＿＿＿＿＿\n"
  }

  var optetoj = docetoj.getBody().setText(encontents);
  optetoj.setFontSize(16);
  optetoj.setFontFamily("Cousine");

  var optjtoe = docjtoe.getBody().setText(jacontents);
  optjtoe.setFontSize(14);
  optjtoe.setFontFamily("Kosugi Maru");

  var docnameetoj = docetoj.getName();
  var docetojid = docetoj.getId();
  var docetojfile = DriveApp.getFileById(docetojid);
  var docetojurl = docetoj.getUrl();
  folder.addFile(docetojfile);
  //DriveApp.getRootFolder().removeFile(docetojfile);

  var docnamejtoe = docjtoe.getName();
  var docjtoeid = docjtoe.getId();
  var docjtoefile = DriveApp.getFileById(docjtoeid);
  var docjtoeurl = docjtoe.getUrl();
  folder.addFile(docjtoefile);

  docetoj.saveAndClose();
  docjtoe.saveAndClose();
  
  //createpdf(folder,docetojid,docnameetoj);
  //createpdf(folder,docjtoeid,docnamejtoe);

  //DriveApp.getRootFolder().removeFile(docjtoefile);

  var docetojBlob = docetojfile.getAs("application/pdf");
  docetojBlob.setName(docnameetoj + ".pdf"); 
  var etojpdf = folder.createFile(docetojBlob);
  etojpdfurl = etojpdf.getUrl();

  var docjtoeBlob = docjtoefile.getAs("application/pdf");
  docjtoeBlob.setName(docnamejtoe + ".pdf");
  var jtoepdf = folder.createFile(docjtoeBlob);
  jtoepdfurl = jtoepdf.getUrl();

  var result = HtmlService.createTemplateFromFile("result");
  result.docetojurl = docetojurl;
  result.docjtoeurl = docjtoeurl;
  result.etojpdfurl = etojpdfurl;
  result.jtoepdfurl = jtoepdfurl;
  result.toppageurl = toppageurl;
  result.grade = e.parameter.grade.toString();
  result.shuffleqs = shuffleqs;
  result.contents = contents;
  result = result.evaluate().setTitle("フォレスタ英単語テスト").setFaviconUrl("-----------------------------------------------");

  return result;
}

/*
function createpdf(folder, docId, filename){
  let url = "https://docs.google.com/document/d/" + docId + "/export?&exportFormat=pdf&format=pdf";
  let token = ScriptApp.getOAuthToken();
  let options = {
    headers: {
      "Authorization": "Bearer" + token
    }
  };
  let blob = UrlFetchApp.fetch(url, options).getBlob().setName(filename + ".pdf");
  folder.createFile(blob);
}
*/

function shuffle(array){
  result = array;
  for(let i = result.length-1; i >= 0; i--){
    var index = Math.floor(Math.random() * i);
    var tmp = result[i];
    result[i] = result[index];
    result[index] = tmp;
  }
  return result;
}
  /*var docurl = "https://docs.google.com/document/d/" + docid + "/export?";
  var docname = doc.getName();
  const opts = {
    exportFormat: 'pdf',      // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       'pdf',      // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         'A4',       // 用紙サイズの指定 legal / letter / A4
    portrait:     'true',     // true → 縦向き、false → 横向き
    fitw:         'true',     // 幅を用紙に合わせるか
    docNames:     'false',    // シート名を PDF 上部に表示するか
    printtitle:   'false',    // スプレッドシート名を PDF 上部に表示するか
    pagenumbers:  'false',    // ページ番号の有無
    gridlines:    'false',    // グリッドラインの表示有無
    fzr:          'false',    // 固定行の表示有無
  };
  const urlExt = [];
  for(optName in opts){
    urlExt.push(optName + '=' + opts[optName]);
  }
  const options  = urlExt.join('&');
  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(docurl + options, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  const blob = response.getBlob().setName(docname + '.pdf');
  var testpdf = folder.createFile(blob);  //　PDFを指定したフォルダに保存
  testpdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);//共有設定をする：「リンクを知っている人」が「閲覧可能」*/
