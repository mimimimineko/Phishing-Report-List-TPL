//https://note.com/skipla/n/n0803937a0887
//https://tetsuooo.net/gas/604/
//https://ja.wikipedia.org/wiki/HTTP%E3%82%B9%E3%83%86%E3%83%BC%E3%82%BF%E3%82%B9%E3%82%B3%E3%83%BC%E3%83%89
//https://service.taksas.tech/?page_id=148
//https://qiita.com/n0bisuke/items/a31a99232e50461eb00f
//https://web.biz-prog.net/gas/gethtml.html#gethtml


//列の指定
//URLを入力する列
let urlColumn                  =  4
//ステータス表示列（OKorNG）
let statusDisplayColumn        =  5
//ステータスコード表示列
let statusCodeDisplayColumn    = 15
//ステータスコード説明列
let statusCodeExpDisplayColumn = 16
//応答速度表示列
let pingDisplay                = 17
//アクセス開始時間表示列
let accessTimeDisplay          = 18


let beforeTime
let afterTime
let ping

function myFunction() {
  const spreadSheetAppName = getSpreadSheetAppObject()
  const spreadSheetName    = getSheetObject(spreadSheetAppName)
  let lastLine             = searchLastLine(spreadSheetName,3)
  for(let i=2;i<=lastLine;i++){
    let url                = getURL(spreadSheetName,i)
    Logger.log(url)
    let statusCode = getstatusCode(url)
    ping           = afterTime - beforeTime
    let statusExp  = statusCodeExp(statusCode)
    outputStatusCode(spreadSheetName,i,statusCode,statusExp)
    outputPing(spreadSheetName,i,ping,beforeTime)
    sendPostTrendmicro(url,i,spreadSheetName)
  }
}




//spreadSheetオブジェクトを取得
function getSpreadSheetAppObject(){
  //変数にはシート名が格納

  //方法1 アクティブなスプレッドシートを取得
  let spreadSheetByActive = SpreadsheetApp.getActive()
  Logger.log(spreadSheetByActive.getName())
  
  //方法2 スプレッドシートのURLを指定し取得
  //let spreadSheetByUrl    = SpreadsheetApp.openByUrl("/*スプレッドシートURL*/")
  //Logger.log(spreadSheetByUrl.getName())
  
  //方法3 スプレッドシートのIDを指定し取得
  //let spreadSheetById     = SpreadsheetApp.openById("/*スプレッドシートID*/")
  //Logger.log(spreadSheetById.getName())


  //方法1 により取得したシート名を返す
  return spreadSheetByActive
}


//sheetObjectを取得
//シート名を受け取る
function getSheetObject(sheetName){
  //2通りで取得

  //方法1 アクティブなシートを取得
  //複数のシートには非対応
  let sheetByActive = sheetName.getActiveSheet()

  //方法2 シート名を指定して取得
  //let sheetByName   = sheetName.getSheetByName(sheetName)
  //Logger.log(sheetByName.getName())
  
  //方法1で取得したシート名を返す。
  return sheetByActive
}

//URL個数check
//上からチェックし、空のセルを検知する
function searchLastLine(sheetName,column){
  let lineValue
  let i=1
  for(i=1;;i++){
    lineValue = sheetName.getRange(i,column)
    //nullの判別はisBlanl()
    if(lineValue.isBlank()==true){
      return i-1
    }
  }
}


function getURL(sheetName,i){
  let url = sheetName.getRange(i,urlColumn)
  return url.getValues()
}

function getstatusCode(url){
  let statusCode
  beforeTime = new Date()
  let ping
  //通常のユーザーエージェント
  try{
    //成功時
    let urlResponse = UrlFetchApp.fetch(
    url,
    {
      muteHttpExceptions: true,
      "validateHttpsCertificates" : false
    })
    statusCode =  urlResponse.getResponseCode()
    afterTime  = new Date()
    return statusCode
  }
  //エラー時
  catch(e){
    Logger.log('DNS error')
    statusCode = 'DNS ERROR'

    //ユーザーエージェントをAndroidに
    try{
      //成功時
      beforeTime = new Date()
      let userAgent = {
      "useragent":'Mozilla/5.0 (Linux; Android 8.0.0; SM-G955U Build/R16NW) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Mobile Safari/537.36'
      }
      let options = {
      "headers" : userAgent
      }
      let urlResponse = UrlFetchApp.fetch(
        url,
        {
          muteHttpExceptions: true,
          "validateHttpsCertificates" : false
        },
        options
      )
      statusCode =  urlResponse.getResponseCode()
      afterTime = new Date()
      return statusCode
    }
    //エラー時
    catch(e){
      Logger.log('DNS error')
      let statusCode = 'DNS ERROR'
      //ユーザエージェントをiPhoneにする
      try{
        //成功時
        beforeTime = new Date()
        let userAgent = {
        "useragent":'Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1'
        }
        let options = {
        "headers" : userAgent
        }
        let urlResponse = UrlFetchApp.fetch(
          url,
          {
            muteHttpExceptions: true,
            "validateHttpsCertificates" : false
          },
          options
        )
        statusCode =  urlResponse.getResponseCode()
        afterTime = new Date()
        return statusCode
      }
      //エラー時
      catch(e){
        beforeTime = new Date()
        afterTime  = new Date()
        Logger.log('DNS error')
        let statusCode = 'DNS ERROR'
        return statusCode
      }
    }
  }
}

//ステータスコードを受け取り、説明を返す
function statusCodeExp(code){
  let exp
  if (code ==100){
    exp = 'Continue' 
  }else if(code == 101){
    exp = 'Switching Protocols'
  }else if(code == 103){
    exp = 'Early Hints'
  }else if(code == 200){
    exp ='OK'
  }else if(code == 201){
    exp = 'Created'
  }else if(code == 202){
    exp = 'Accepted'
  }else if(code == 203){
    exp = 'Non-Authoritative Information'
  }else if(code == 204){
    exp = 'No Content'
  }else if(code == 205){
    exp ='Reset Content'
  }else if(code == 206){
    exp = 'Partial Content'
  }else if(code == 300){
    exp = 'Multiple Choices'
  }else if(code == 301){
    exp = 'Moved Permanently'
  }else if(code == 302){
    exp = 'Found'
  }else if(code == 303){
    exp ='See Other'
  }else if(code == 304){
    exp = 'Not Modified'
  }else if(code == 307){
    exp = 'Temporary Redirect'
  }else if(code == 308){
    exp = 'Permanent Redirect'
  }else if(code == 400){
    exp = 'Bad Request'
  }else if(code == 401){
    exp ='Unauthorized'
  }else if(code == 402){
    exp = 'Payment Required'
  }else if(code == 403){
    exp = 'Forbidden'
  }else if(code == 404){
    exp ='Not Found'
  }else if(code == 405){
    exp = 'Method Not Allowed'
  }else if(code == 406){
    exp = 'Not Acceptable'
  }else if(code == 407){
    exp = 'Proxy Authentication Required'
  }else if(code == 408){
    exp = 'Request Timeout'
  }else if(code == 409){
    exp ='Conflict'
  }else if(code == 410){
    exp = 'Gone'
  }else if(code == 411){
    exp = 'Length Required'
  }else if(code == 412){
    exp = 'Precondition Failed'
  }else if(code == 413){
    exp = 'Payload Too Large'
  }else if(code == 414){
    exp ='URI Too Long'
  }else if(code == 415){
    exp = 'Unsupported Media Type'
  }else if(code == 416){
    exp = 'Range Not Satisfiable'
  }else if(code == 417){
    exp = 'Expectation Failed'
  }else if(code == 418){
    exp = "I'm a teapot"
  }else if(code == 422){
    exp ='Unprocessable Entity'
  }else if(code == 425){
    exp = 'Too Early'
  }else if(code == 426){
    exp = 'Upgrade Required'
  }else if(code == 428){
    exp = 'Precondition Required'
  }else if(code == 429){
    exp = 'Too Many Requests'
  }else if(code == 431){
    exp = 'Request Header Fields Too Large'
  }else if(code == 451){
    exp ='Unavailable For Legal Reasons'
  }else if(code == 500){
    exp = 'Internal Server Error'
  }else if(code == 501){
    exp = 'Not Implemented'
  }else if(code == 502){
    exp = 'Bad Gateway'
  }else if(code == 503){
    exp = 'Service Unavailable'
  }else if(code == 504){
    exp ='Gateway Timeout'
  }else if(code == 505){
    exp = 'HTTP Version Not Supported'
  }else if(code == 506){
    exp = 'Variant Also Negotiates'
  }else if(code == 507){
    exp = 'Loop Detected'
  }else if(code == 510){
    exp ='Not Extended'
  }else if(code == 511){
    exp = 'Network Authentication Required'
  }else{
    exp = 'Not Defined'
  }
  return exp
}

//ステータスコード表示
function outputStatusCode(sheetName,i,code,status){
  let outputCell = sheetName.getRange(i,statusDisplayColumn)
  outputCell.setValue(code)
  let outputOkOrNgCell =sheetName.getRange(i,statusCodeDisplayColumn)
  let outputStatusCell = sheetName.getRange(i,statusCodeExpDisplayColumn)
  outputStatusCell.setValue(status)

  if (code ==200){
    //通常OK
    outputOkOrNgCell.setValue('OK')
    outputOkOrNgCell.setBackground('#00ff00')
    outputCell.setBackground('#00ff00')
    outputStatusCell.setBackground('#00ff00')
  }else if(code == 403){
    //認証拒否 ロボット検知など
    outputOkOrNgCell.setValue('OK')
    outputOkOrNgCell.setBackground('#808000')
    outputCell.setBackground('#808000')
    outputStatusCell.setBackground('#808000')
  }else if(code == 404){
    //ファイル削除されて存在しない　ユーザエージェントによって送信拒否されるとき
    outputOkOrNgCell.setValue('NG(UAによってはOKの可能性あり)')
    outputOkOrNgCell.setBackground('#808000')
    outputCell.setBackground('#808000')
    outputStatusCell.setBackground('#808000')
  }else{
    outputOkOrNgCell.setValue('NG')
    outputOkOrNgCell.setBackground('#ff0000')
    outputCell.setBackground('#ff0000')
    outputStatusCell.setBackground('#ff0000')
  }
}

function outputPing(sheetName,i,ping,accessTime){
  let outputCell = sheetName.getRange(i,pingDisplay)
  outputCell.setValue(ping + ' ms')
  let outputAccessTimeCell = sheetName.getRange(i,accessTimeDisplay)
  let formatedAccessTime = "JST - "+Utilities.formatDate(accessTime, "JST", "yyyy/MM/dd (E) HH:mm:ss Z")
  outputAccessTimeCell.setValue(formatedAccessTime)
}

function sendPostTrendmicro(url){
  //トレンドマイクロの結果表示ページ（PHP）はPOSTでURLを受け取る
  var payload =
   {
     "message" : url
   }
  let options = {
    "method"  : "post",
    "payload" : payload,
  }
  let html = UrlFetchApp.fetch("https://global.sitesafety.trendmicro.com/result.php",options)
  let result = Parser.data(html).from('<div class="labeltitleresult">').to('</div>')
}
