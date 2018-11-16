// user configuration values
var default_destaddress = ["hoge@fuga.com","foo@bar.com"];
var default_url_spreadsheet = "";

// system configuration values
var default_sheet_name = "current"; // the name of a sheet from which Order list is created.

var default_rowindex_first    = 2;
var default_colindex_first    = 1;
var default_colindex_last     = 8;

var default_colindex_day      = 1;
var default_colindex_partname = 2;
var default_colindex_number   = 3;
var default_colindex_orddest  = 4;
var default_colindex_url      = 5;
var default_colindex_person   = 6;
var default_colindex_orderday = 7;
var default_colindex_extra    = 8;

/**
 *  OrderListNotification
 *
 *  部品係が注文する物品のリストをメールで送る
 *
 */
function OrderListNotification( dest_address ) {
  if ( destaddress === undefined ) {
      var destAddress = "";
      for ( var index in list_address ) {
        destAddress += list_address[index] + "," ;
      }
  }

  var date     = new Date();
  var mailbody = createMailBody();
  var title = "[部品係]"
                + date.getFullYear()  + "年 "
                + (date.getMonth()+1) + "月 "
                + (date.getDate())    + "日 "
                + "部品係発注リスト通知";
  MailApp.sendEmail( destAddress, title, mailbody );
}

/**
 *  createMailBody
 *
 *  getOrderListで得られた発注リストからメール本文を作成
 *
 *  @return {string} mailbody (メール本文)
 */
function createMailBody()
{
    var orderlist = getOrderList( default_url_spreadsheet );

    // URLでソート
    orderlist.sort(function(a,b){return a[default_colindex_url-1] < b[default_colindex_url-1];});

    if ( orderlist.length == 0 ) {
        var mailbody =
            "== 今週の購入部品 ==\n"
          + "URL:" default_url_spreadsheet + "\n\n";
        for ( var i=0,l=orderlist.length; i<l; i++ ) {
            mailbody += 
                value[i][default_colindex_partname-1] + "\n"
              + "  " + "記入日:" + value[i][default_colindex_day-1]     + "\n"
              + "  " + "個数  :" + value[i][default_colindex_number-1]  + "\n"
              + "  " + "発注先:" + value[i][default_colindex_orddest-1] + "\n"
              + "  " + "URL   :" + value[i][default_colindex_url-1]     + "\n"
              + "  " + "記入者:" + value[i][default_colindex_person-1]  + "\n"
              + "  " + "備考  :" + value[i][default_colindex_extra-1]   + "\n"
              + "\n";
        }
    } else {
        var mailbody = "今週の発注部品はありません.";
    }

    return mailbody;
}

/**
 * getOrderList
 *
 * スプレッドシートのURLから,発注する部品のリストを生成する
 *
 * @param {string} url (スプレッドシートのURL)
 * @return {Array} orderlist (部品リスト)
 * @throws {Error} cannot open sheet (シートが開けない)
 */
function getOrderList( url )
{
    if ( url === undefined ) {
        url = default_url_spreadsheet;
    }

    var ss    = SpreadsheetApp.openByUrl( url );
    var sheet = ss.getSheetByName( default_sheet_name );
    if ( ss == null || sheet == null ) {
        throw new Error( "cannot open sheet" );
    }

    var rowindex_first = default_rowindex_first;
    var rowindex_last  = sheet.getLastRow();
    var colindex_first = default_colindex_first;
    var colindex_last  = default_colindex_last;

    var colindex_orderday = default_colindex_orderday;

    var values = sheet.getRange( rowindex_first
                                ,colindex_first
                                ,rowindex_last - rowindex_first + 1
                                ,colindex_last - colindex_first + 1 ).getValues();

    var orderlist = [];

    for ( var i=0; i<rowindex_last - rowindex_first; i++ ) {
        // 記入日が空白でなく,購入日が空白の行をArrayに追加
        if ( values[i][colindex_orderday-1] == "" && values[i][0] != "" ) {
            orderlist.push(values[i]);
        }
    }

    return orderlist;
}

