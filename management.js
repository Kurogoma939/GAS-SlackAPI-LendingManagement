/**********************
 * 　　　共通設定
 *********************/
// スプレッドシートから値を取得する
let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let data = spreadsheet.getSheetByName("貸出状況");
let master = spreadsheet.getSheetByName("マスターシート");
const token = "xoxb-1452582367283-xxxxXXXXxxxxXXXX-xxxxXXXXxxxxXXXX";
const slackApp = SlackApp.create(token);

/*********************
 * 定期メッセージ
 *********************/
//Slackに投稿するメッセージを定義する
let monthly_message = `
  <!channel>
  おはようございます！
  オリエントの棚に追加する本の募集をいたします！

  購入して欲しい本がある方は、以下のアンケートに回答願います。

  それでは、よろしくお願いします！
  https://forms.gle/6bwNCHWTHFSu6c8G9

`;
// 実行する関数
function survey() {
	//Slackボットがメッセージを投稿するチャンネルを定義する
	let channelId = "C02N59G9K4P"; // 図書委員会
	let options = {};

	var access = slackApp.postMessage(channelId, monthly_message, options);
	Logger.log(access);
}

function onEdit(e) {
	// D列を選択したら、E列に「当日の日付」、F列に「20日後の日付」を自動入力
	let rng = data.getActiveCell();
	let row = e.range.getRow();
	//日付取得
	const date = new Date();
	const returnDate = new Date(date.setDate(date.getDate() + 20));

	if (rng.getColumn() == 4) {
		sheet.getRange('E' + row.toString()).setValue(new Date());
		sheet.getRange('F' + row.toString()).setValue(returnDate);
	}
}

function msgBox() {
	// G列のチェックボックスを押したらダイアログを表示させOKを押したらD~Gをリセットする
	let rng = data.getActiveCell();
	let currentRow = rng.getRow();
	const flag = rng.getValue();

	if (rng.getColumn() == 7 && flag == true) {
		const msg = "返却されましたか？";
		const answer = Browser.msgBox(msg, Browser.Buttons.OK_CANCEL);
		if (answer == 'ok') {
			sheet.getRange('D' + currentRow).setValue("オフィス");
			sheet.getRange('E' + currentRow).clearContent();
			sheet.getRange('F' + currentRow).clearContent();
			sheet.getRange('G' + currentRow).uncheck();
		} else {
			sheet.getRange('G' + currentRow).uncheck();
		}
	}
}

function notification() {
	function formatDate(dt) {
		var y = dt.getFullYear();
		var m = ('00' + (dt.getMonth() + 1)).slice(-2);
		var d = ('00' + dt.getDate()).slice(-2);
		return (y + '-' + m + '-' + d);
	}

	function getDateDiff(dateString1, dateString2) {
		var date1 = new Date(dateString1);
		var date2 = new Date(dateString2);
		var msDiff = date1.getTime() - date2.getTime();
		return Math.ceil(msDiff / (1000 * 60 * 60 * 24));
	}

	let today = new Date;
	let todayDate = formatDate(today);
	// データ処理のための配列
	let arrayName = [];
	let arrayId = [];
	let arrayTitle = [];
	let arrayDate = [];

	let sheet = SpreadsheetApp.getActive().getSheetByName("貸出状況");
	for (let i = 2; i < sheet.getLastRow(); i++) {
		let range = sheet.getRange("F" + i);
		let isDone = sheet.getRange("G" + i).getValue();

		if (!range.isBlank()) {
			let origSpreadDate = range.getValue();
			let spreadDate = formatDate(origSpreadDate);

			if (getDateDiff(todayDate, spreadDate) >= 0 && !isDone) {
				// 貸出者を取得
				let nameExpired = sheet.getRange("D" + i).getValue();
				arrayName.push(nameExpired);
				// 書籍名を取得
				let titleExpired = sheet.getRange("B" + i).getValue();
				arrayTitle.push(titleExpired);
				// F列の日付を取得
				let dateFormed = formatDate(sheet.getRange("F" + i).getValue());
				arrayDate.push(dateFormed);

				if (getDateDiff(todayDate, spreadDate) == 0) {
					Logger.log('今日が返却期限です。');
				} else {
					Logger.log('返却期限が' + getDateDiff(todayDate, spreadDate) + '日過ぎています。');
				}

				let dataList = master.getRange(2, 1, 40, 2).getValues();

				for (let j = 0; j < 40; j++) {
					if (dataList[j][0] == nameExpired) {
						arrayId.push(dataList[j][1]);
					}
				}
			}
		}
	}

	// 期限がすぎた人の名前
	Logger.log(arrayName);
	// それを使って取ったid
	Logger.log(arrayId);
	// 本のタイトル
	Logger.log(arrayTitle);
	// 返却日
	Logger.log(arrayDate);

	for (let k = 0; k < arrayName.length; k++) {
		let librarian_message = `
    お疲れさまです！図書委員会です。
    貸出中の本が返却されていません。
    タイトル：${arrayTitle[k]}
    返却日　：${arrayDate[k]}
    ご確認をお願いします。
    `;

		Logger.log(librarian_message);

		let channelId = '';

		if (!(arrayId[k] == '' || arrayId[k] == null)) {
			channelId = arrayId[k]; // 本番環境
			// let channelId = "U01SWRQJVSN"; // 宛先をテスト用に佐生にしている
			// let channelId = "";
		} else {
			channelId = "C02N59G9K4P"; // 図書委員会チャンネル
		}

		let options = {
			"as_user": true, // トークン発行者の名前が設定できる
		}

		let access = slackApp.postMessage(channelId, librarian_message, options);
		Logger.log(access);
	}
}
