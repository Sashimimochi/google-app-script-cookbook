// 今月の情報を取得する
function getThisMonth() {
    var today = new Date();
    var thisMonth = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMM');
    return thisMonth;
}

function getSpreadSheet(sheetName) {
    var book = SpreadsheetApp.getActiveSpreadsheet(); //起動中のスプレッドシートを選択する
    var sheet = book.getSheetByName(sheetName); //シートを選択する
    return sheet
}

function writeSpreadSheet(events) {
    //取得したデータをスプレッドシートに書き込む
    var sheet = getSpreadSheet("sheet1"); //シートを選択する
    var rows = events.length; //書き込むデータのサイズを取得する
    var cols = events[0].length; //書き込みデータのカラム数を取得する
    var lastRow = sheet.getLastRow(); //スプレッドシートに書き込まれている最終行番号を取得する
    var colStartIndex = 1; //開始行番号
    sheet.getRange(lastRow + 1, colStartIndex, rows, cols).setValues(events); //選択した範囲にデータを書き込む
}

// イベントデータを取得する
function registerEventData() {
    var baseURL = '{イベントデータを取得する先のAPIのURL}'
    var thisMonth = getThisMonth();
    var url = baseURL + '?ym=' + thisMonth; //必要に応じてパラメータを指定
    var response = UrlFetchApp.fetch(url).getContentText(); //APIを叩いてデータを取得する
    var json = JSON.parse(response); //json形式にパース

    var jsonLength = Object.keys(json.events).length; //取得したjsonデータの長さを取得(ヒット件数)
    var events = []; //スプレットシートに書き込むデータを格納するための配列を確保

    if (jsonLength < 1) { //ヒット件数が0件なら何もしない
        return
    }

    var eventIds = getEventIds(); //登録済みのイベントデータのID一覧を取得する

    for (var i = 0; i < jsonLength; i++) {
        var overlap = checkRegisteredId(json.events[i].event_id, eventIds); //すでに登録済みのイベントかどうか
        if (overlap) { //未登録のデータであれば書き込み用配列に追加
            events.push([
                json.events[i].event_id //jsonデータから欲しいキーのデータを取得する
            ]);
        }
    }

    writeSpreadSheet(events);
    Utilities.sleep(2000); //APIサーバに負荷を掛けないように一定時間空ける
}

// 取得済みのイベントかどうかを判定する
function checkRegisteredId(targetId, eventIds) {
    var overlap = eventIds.some(function (array, i, eventIds) {
        return (array[0] === targetId); // eventIdsの中にtargetIdがあればreturn
    });

    if (overlap) { // すでに登録済みのイベントかどうか
        return false;
    } else {
        return true;
    }
}

// イベントIDの一覧を取得
function getEventIds() {
    var sheet = getSpreadSheet("sheet1"); //シートを選択する
    var lastRow = sheet.getLastRow(); //最終行の行番号を取得する

    var ids = sheet.getRange(2, 1, lastRow - 1).getValues(); //選択した範囲のイベントIDを取得する

    return ids;
}

// スプレッドシートからイベント情報を取得する
function getEventData() {
    var sheet = getSpreadSheet("sheet1"); //シートを選択する
    var firstRange = sheet.getRange(1, 1, 1, 9); //キーとなるセルを選択する
    var firstRowValues = firstRange.getValues(); //選択したセルの値を取得する
    var titleColumns = firstRowValues[0]; //キーとなるカラムを指定する
    var lastRow = sheet.getLastRow(); //最終行の行番号を取得する
    var rowValues = []; //データ格納用の配列を確保する

    for (var rowIndex = 2; rowIndex <= lastRow; rowIndex++) { //スプレッドシートに書き込まれているイベントを全件取得する
        var colStartIndex = 1; //開始列番号
        var rowNum = 1; //一度に取得する行数
        var range = sheet.getRange(rowIndex, colStartIndex, rowNum, sheet.getLastColumn()); //指定した範囲のセルを取得する
        var values = range.getValues(); //指定したセルの値を取得する
        rowValues.push(values[0]); //データを配列に格納する
    }

    var jsonArray = []; //データ格納用の配列を確保する
    for (var i = 0; i < rowValues.length; i++) {
        var line = rowValues[i]; //1行ずつデータを読み込む
        var json = new Object(); //空のオブジェクトを生成する
        if (!isFinishedDate(line[4])) { //イベント終了日が現在時刻より後のイベントのみを取得する
            for (var j = 0; j < titleColumns.length; j++) {
                json[titleColumns[j]] = line[j];
            }
            jsonArray.push(json) //配列にデータを格納する
        }
    }

    return jsonArray
}

// targetDateが現在時刻より古いか判定する
function isFinishedDate(targetDate) {
    var today = new Date(); //現在時刻を取得する
    var isFinished = Moment.moment(targetDate).isBefore(today); //Moment.jsを使って日付の比較をする(boolean)
    return isFinished
}

// FireStoreの情報を取得する
// Firebaseのコンソール画面から取得した秘密鍵の情報を記述する
function getFirestoreCertification() {
    var certification = {
        'email': '{client_emailの値}',
        'key': '-----BEGIN PRIVATE KEY-----\n{private_keyの値}\n-----END PRIVATE KEY-----\n',
        'projectId': '{project_idの値}'
    }
    return certification
}

// FireStoreのコレクションを新規作成して、データを書き込む
function createFirestore() {
    var eventData = getEventData(); //書き込むためのイベントデータを取得
    var certification = getFirestoreCertification(); //FireStoreの認証情報を取得
    var firestore = FirestoreApp.getFirestore(certification.email, certification.key, certification.projectId); //FireStoreに接続
    firestore.createDocument('EventCollection/Eventdata', eventData); //Firestoreのコレクションを新規作成して、データを書き込み
}

// 作成済みのFirestoreのコレクションにデータを書き込む
function updateFirestore() {
    var eventData = getEventData();
    var certification = getFirestoreCertification();
    var firestore = FirestoreApp.getFirestore(certification.email, certification.key, certification.projectId);
    firestore.updateDocument('EventCollections/EventData', eventData); //作成済みのコレクションにデータを書き込み
}