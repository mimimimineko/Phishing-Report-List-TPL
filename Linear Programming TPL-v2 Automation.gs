// function getSpreadSheetAppObject()                     // spreadSheetオブジェクトを取得
// function getSheetObject(sheetName)                     // sheetObjectを取得 シート名を受け取る
// function lastColumn(sheetName,i,p)                       // i行がp列目からどこまで埋まっているかを返す
// function lastLine(sheetName,i,p)                         // i列がp行目からどこまで埋まっているかを返す
// function setKoumoku1(sheetName,setline,hensuuKazu,hutoushikiKazu,opLine)           //最初の設定、2行目にX4，X5，X6、、をセットする
// function setKoumoku2(sheetName,setColumn,hensuuKazu,hutoushikiKazu,kaishiLine){    //最初の設定、2列目にX4，X5，X6、、をセットする
// function copyLine(sheetName,copySakiLine,copyMotoLine,kaishiColumn,owariColumn)    //行をコピーする
// function copyColumn(sheetName,copySakiColumn,copyMotoColumn,kaishiLine,owariLine)  //列をコピーする
// function setKoumoku3(sheetName,hidariueHashiLine,hidariueHashiColumn,migishitaHashiLine,migishitaHashiColumn)        //10をセットする
// function step01(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu)                                                 // 線形計画法第三回講義パワポSTEP1　38枚目
// function checkLineHunokazu(sheetName,line,p,q)                                  // line行にp列目からq列目の範囲で負の数があるかをチェックあれば1を返却、無ければ0を返却
// function step02(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu)                                                 // 線形計画法第三回講義パワポSTEP1　41枚目
// function serchMinInLine(sheetName,line,p,q)        line行のp列からq列までで最小値を探し最小値valueと列columnを返す
// function step03to4(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu)                                                 // 線形計画法第三回講義パワポSTEP1　42枚目

function myFunction() {
  // シート設定
  let sheetAppObjName =    getSpreadSheetAppObject();
  let sheetName       =    getSheetObject(sheetAppObjName);

  // 開始行列設定
  let opLine      =2;   // 2行目までが項目
  let opColumn    =2;   // B列　までが項目

  // 不等式の数を数える
  let hutoushikiKazu    = lastLine(sheetName,opLine+1,opLine+1);
  console.log("不等式の数　：　"+hutoushikiKazu);
  
  // 変数の数を数える
  let hensuuKazu        = lastColumn(sheetName,opColumn+1,opColumn+1);
  console.log("　変数の数　：　"+hensuuKazu);

  // 線形計画法第三回講義パワポSTEP1　38枚目
  step01(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu);

  // 線形計画法第三回講義パワポSTEP1　41枚目
  for(let continueOrEnd=1;continueOrEnd!=0;){
    continueOrEnd=step02(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu);

    // 線形計画法第三回講義パワポSTEP1　42枚目

    let pibotData=step03to4(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu);
    console.log("pibotData :  "+pibotData.pibot +"  "+pibotData.pibotLine+"行　" + pibotData.pibotColumn +"列　");

    step5(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu,pibotData.pibot,pibotData.pibotLine,pibotData.pibotColumn);

    opLine=opLine+hutoushikiKazu+2;

    continueOrEnd=step02(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu);
  }
}
function step5(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu,pibot,pibotLine,pibotColumn){
  console.log(opLine+hutoushikiKazu);
  console.log("次の表の数値が入る部分は"+(opLine+hutoushikiKazu+3)+" 行から　"+(opLine+hutoushikiKazu+3+hutoushikiKazu)+"行までの間に初期値セット");
  for(let i=opColumn+1;i<=opColumn+hensuuKazu+hutoushikiKazu+1;i++){
    if(i!=opColumn+hensuuKazu+hutoushikiKazu+1){
      sheetName.getRange(opLine+hutoushikiKazu+2,i).setValue("X"+(i-opColumn));
    }else{
      sheetName.getRange(opLine+hutoushikiKazu+2,i).setValue("基底ｂ");
    }
  }
  for(let i=opColumn+1;i<=opColumn+hensuuKazu+hutoushikiKazu+1;i++){
    sheetName.getRange(pibotLine+hutoushikiKazu+2,i).setValue(sheetName.getRange(pibotLine,i).getValue()/pibot);
  }
  console.log("step5  "+(pibotLine+hutoushikiKazu+2)+"行");

  for(let i=opLine+hutoushikiKazu+3;i<=opLine+hutoushikiKazu+3+hutoushikiKazu;i++){
    sheetName.getRange(i,opColumn).setValue(sheetName.getRange(i-hutoushikiKazu-2,opColumn).getValue());
    if(i==pibotLine+hutoushikiKazu+2){
      console.log(i+"行飛ばした");
      continue;
    }
    let valueA=sheetName.getRange(i-hutoushikiKazu-2,pibotColumn).getValue(); //引く値の 列ごとに掛ける 上の表の各値
    for(let j=opColumn+1;j<=opColumn+hutoushikiKazu+hensuuKazu+1;j++){
      // let valueB=sheetName.getRange(opLine+hutoushikiKazu+3,j).getValue();
      let valueB=sheetName.getRange(pibotLine+hutoushikiKazu+2,j).getValue();
      let valueC=sheetName.getRange(i-hutoushikiKazu-2,j).getValue(); //上の表の引かれる各値
      sheetName.getRange(i,j).setValue(valueC-valueA*valueB);
      console.log((valueC-valueA*valueB)+" A："+valueA+"("+(i-hutoushikiKazu-2)+", "+pibotColumn+")"+"  B:"+valueB+"("+(opLine+hutoushikiKazu+3)+", "+j+")"+"  C:"+valueC+"("+(i-hutoushikiKazu-2)+", "+j+")");
      console.log(i,j);
    }
    console.log(i+" 行終わり");
  }
  sheetName.getRange(pibotLine+hutoushikiKazu+2,opColumn).setValue((sheetName.getRange(opLine,pibotColumn).getValue()));

}
function step03to4(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu){
  // STEP3
  let searchLine=lastLine(sheetName,opColumn+1,1);
  let step3minValueCellData=serchMinInLine(sheetName,searchLine,opColumn+1,opColumn+hensuuKazu+hutoushikiKazu+1);

  // STEP4
  let columnA       =step3minValueCellData.column;
  let columnB       =opColumn+hensuuKazu+hutoushikiKazu+1;
  let columnBwaruA  =columnB+1;

  for(let i=searchLine-hutoushikiKazu;i<searchLine;i++){
    let value=(sheetName.getRange(i,columnB).getValue())/(sheetName.getRange(i,columnA).getValue());
    sheetName.getRange(i,columnBwaruA).setValue(value);
    console.log("function step03to4 :  "+i+"行"+columnB+     "列" +"   値： "+sheetName.getRange(i,columnB).getValue());
    console.log("function step03to4 :  "+i+"行"+columnA+     "列" +"   値： "+sheetName.getRange(i,columnA).getValue());
    console.log("function step03to4 :  "+i+"行"+columnBwaruA+"列" +"   値: " +value);
  }

  let step4minValueCellData=serchMinInColumnSeinosuNiKagiru(sheetName,columnBwaruA,opLine+1,opLine+hutoushikiKazu);


  // STEP5
  let pibot=sheetName.getRange(step4minValueCellData.line,step3minValueCellData.column).getValue();
  return {
    pibot:pibot,
    pibotColumn:columnA,
    pibotLine:step4minValueCellData.line
  };
}
function serchMinInColumn(sheetName,column,p,q){
  let minValue;
  let minCell=p;
  for(let a=p;a<=q;a++){
    let value=sheetName.getRange(a,column).getValue();
    if(a==p){
      minValue=value;
    }
    if(value<minValue){
      minValue=value;
      minCell=a;
    }
  }
  console.log("function serchMinInColumn "+column + "　列の 最小値は　"+minCell+"　行の　"+minValue+"　です！　");
  console.log("function serchMinInColumn return"+{value:minValue,column:minCell});
  return{
    value:minValue,
    line:minCell
  }
}
function serchMinInColumnSeinosuNiKagiru(sheetName,column,p,q){
  let minValue;
  let minCell=p;
  for(let a=p;a<=q;a++){
    let value=sheetName.getRange(a,column).getValue();
    if(value>=0){
      if(a==p){
        minValue=value;
      }else if(minValue==undefined||minValue==null){
        minValue=value;
      }
      if(value<minValue){
        minValue=value;
        minCell=a;
      }
    }
  }
  console.log("function serchMinInColumnSeinosuNiKagiru "+column + "　列の 最小値は　"+minCell+"　行の　"+minValue+"　です！　");
  console.log("function serchMinInColumnSeinosuNiKagiru return"+{value:minValue,column:minCell});
  return{
    value:minValue,
    line:minCell
  }
}
function serchMinInLine(sheetName,line,p,q){
  let minValue;
  let minCell=p;
  for(let a=p;a<=q;a++){
    let value=sheetName.getRange(line,a).getValue();
    if(a==p){
      minValue=value;
    }
    if(value<minValue){
      minValue=value;
      minCell=a;
    }
  }
  console.log("function serchMinInLine "+line + "　行の 最小値は　"+minCell+"　列の　"+minValue+"　です！　");
  console.log("function serchMinInLine return"+{value:minValue,column:minCell});
  return{
    value:minValue,
    column:minCell
  }
}
function step02(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu){
  let searchLine=lastLine(sheetName,opColumn+1,1);
  let checkResultHunokazu=checkLineHunokazu(sheetName,searchLine,opColumn+1,opColumn+hutoushikiKazu+hensuuKazu+1);
  console.log("function step02 " +"負の数あり→1　負の数なし→0　　：　"+checkResultHunokazu+"   "+searchLine+"行目");
  return checkResultHunokazu;
}
function checkLineHunokazu(sheetName,line,p,q){
  let a;
  for(a=p;a<=q;a++){
    if(sheetName.getRange(line,a).getValue()<0){
      console.log("function checkLineHunokazu "+1);
      return 1;
    }
  }
  console.log("function checkLineHunokazu "+0);
  return 0;
}
function step01(sheetName,opLine,opColumn,hutoushikiKazu,hensuuKazu){
  // 不等式の数に応じてopLine行にX4，X5、X6、、、をセット
  setKoumoku1(sheetName,opLine,hensuuKazu,hutoushikiKazu,opLine+1);

  // 変数の数に応じてopColumn列にX4，X5、X6、、、をセット
  setKoumoku2(sheetName,opColumn,hensuuKazu,hutoushikiKazu,opLine+1);

  // fをセット
  copyLine(sheetName,opLine+hutoushikiKazu+1,1,opColumn,20);

  // bをセット
  copyColumn(sheetName,opColumn+hutoushikiKazu+hensuuKazu+1,1,2,20);

  // 10をセット
  setKoumoku3(sheetName,opLine+1,opColumn+hensuuKazu+1,opLine+hutoushikiKazu,opColumn+hutoushikiKazu+hensuuKazu);
}
function setKoumoku3(sheetName,hidariueHashiLine,hidariueHashiColumn,migishitaHashiLine,migishitaHashiColumn){
  for(let n=0;n<migishitaHashiColumn-hidariueHashiColumn+1;n++){
    for(let i=0;i<migishitaHashiLine-hidariueHashiLine+1;i++){
      if(i==n){
        sheetName.getRange(hidariueHashiLine+i,hidariueHashiColumn+n).setValue(1);
      }else{
        sheetName.getRange(hidariueHashiLine+i,hidariueHashiColumn+n).setValue(0);
      }
    }
  }
  for(let j=0;j<=migishitaHashiLine-hidariueHashiLine+1;j++){
    sheetName.getRange(migishitaHashiLine+1,hidariueHashiColumn+j).setValue(0);
  }
}
function copyColumn(sheetName,copySakiColumn,copyMotoColumn,kaishiLine,owariLine){
  for(let n=kaishiLine;n<=owariLine;n++){
    sheetName.getRange(n,copyMotoColumn).copyTo(sheetName.getRange(n,copySakiColumn));
  }
}
function copyLine(sheetName,copySakiLine,copyMotoLine,kaishiColumn,owariColumn){
  for(let n=kaishiColumn;n<=owariColumn;n++){
    sheetName.getRange(copyMotoLine,n).copyTo(sheetName.getRange(copySakiLine,n));
  }
}
function setKoumoku2(sheetName,setColumn,hensuuKazu,hutoushikiKazu,kaishiLine){
  for(let n=hensuuKazu+1;n<=hensuuKazu+hutoushikiKazu;n++){
    sheetName.getRange(kaishiLine+n-hensuuKazu-1,setColumn).setValue("X"+n);
  }
}
function setKoumoku1(sheetName,setline,hensuuKazu,hutoushikiKazu,kaishiColumn){
  for(let n=1;n<=hensuuKazu+hutoushikiKazu;n++){
    sheetName.getRange(setline,n+kaishiColumn-1).setValue("X"+n);
  }
}
function lastLine(sheetName,i,p){
  let n;
  for(n=p;  sheetName.getRange(n,i).isBlank()!=true;  n++){}
  return n-p;
}
function lastColumn(sheetName,i,p){
  let n;
  for(n=p;  sheetName.getRange(i,n).isBlank()!=true;  n++){}
  return n-p;
}
function getSheetObject(sheetName){
  //2通りで取得

  //方法1 アクティブなシートを取得
  //複数のシートには非対応
  let sheetByActive = sheetName.getActiveSheet();

  //方法2 シート名を指定して取得
  //let sheetByName   = sheetName.getSheetByName(sheetName);
  
  //方法1で取得したシート名を返す。
  return sheetByActive;
}
function getSpreadSheetAppObject(){
  //変数にはシート名が格納
  //方法1 アクティブなスプレッドシートを取得
  let spreadSheetByActive = SpreadsheetApp.getActive();
  Logger.log(spreadSheetByActive.getName());
  
  //方法2 スプレッドシートのURLを指定し取得
  //let spreadSheetByUrl    = SpreadsheetApp.openByUrl("/*スプレッドシートURL*/");
  //Logger.log(spreadSheetByUrl.getName());
  
  //方法3 スプレッドシートのIDを指定し取得
  //let spreadSheetById     = SpreadsheetApp.openById("/*スプレッドシートID*/");
  //Logger.log(spreadSheetById.getName());

  //方法1 により取得したシート名を返す
  return spreadSheetByActive;
}
