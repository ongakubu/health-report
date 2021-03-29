/*各ファイルのIDの定義*/
/*例えばスプレッドシートを開いたときのURLにおいては、https://docs.google.com/spreadsheets/d/ここの長い文字列がID/edit*/
var registerSheetId   = '登録フォーム（回答）スプレッドシートのID';
var recipientsSheetId = '宛先一覧スプレッドシートのID';

/*登録フォームからメールアドレスを読み込んでrecipientsへコピーする（健康観察開始前）*/
function getRecipients () {
  var registerSheet     = SpreadsheetApp.openById(registerSheetId).getActiveSheet();      //register_form_answerのID
  var nRecipients       = registerSheet.getLastRow()-1;                                   // 登録者の人数（=行数-1）
  var recipientsSheet   = SpreadsheetApp.openById(recipientsSheetId).getActiveSheet();    // recipientsを読み込む
  recipientsSheet.clear();
  
  for (var i=1; i<=nRecipients; i++){
    recipientsSheet.getRange(1, 1, 1, 1).setValue('学籍番号');       // 1行目（見出し行）に「学籍番号」と入力する
    recipientsSheet.getRange(1, 2, 1, 1).setValue('パート');        // 1行目（見出し行）に「氏名」と入力する
    recipientsSheet.getRange(1, 3, 1, 1).setValue('氏名');          // 1行目（見出し行）に「メールアドレス」と入力する
    recipientsSheet.getRange(1, 4, 1, 1).setValue('メールアドレス');  // 1行目（見出し行）に「メールアドレス」と入力する
    var studentno = registerSheet.getRange(i+1, 2).getValue();     // register_form_answerから学籍番号を取得する
    recipientsSheet.getRange(i+1, 1).setValue(studentno);          // recipientsに学籍番号を入力する
    var part      = registerSheet.getRange(i+1, 3).getValue();     // register_form_answerからパートを取得する
    recipientsSheet.getRange(i+1, 2).setValue(part);               // recipientsにパートを入力する
    var name      = registerSheet.getRange(i+1, 4).getValue();     // register_form_answerから氏名を取得する
    recipientsSheet.getRange(i+1, 3).setValue(name);               // recipientsに氏名を入力する
    var mailto    = registerSheet.getRange(i+1, 5).getValue();     // register_form_answerからメールアドレスを取得する
    recipientsSheet.getRange(i+1, 4).setValue(mailto);             // recipientsにメールアドレスを入力する
  }
}

/*登録フォームからメールアドレスを読み込んでrecipientsへコピーする（健康観察開始後）*/
/*健康観察期間の途中で登録したユーザを自動追加する */
function addRecipients () {
  var registerSheet     = SpreadsheetApp.openById(registerSheetId).getActiveSheet();    //register_form_answerのID
  var nRecipients       = registerSheet.getLastRow()-1;                                 // 登録者の人数（=行数-1）
  var recipientsSheet   = SpreadsheetApp.openById(recipientsSheetId).getActiveSheet();  // recipientsを読み込む
  
  for (var i=1; i<=nRecipients; i++){
    var studentno = registerSheet.getRange(i+1, 2).getValue();  // register_form_answerから学籍番号を取得する
    recipientsSheet.getRange(i+1, 1).setValue(studentno);       // recipientsに学籍番号を入力する
    var part      = registerSheet.getRange(i+1, 3).getValue();  // register_form_answerからパートを取得する
    recipientsSheet.getRange(i+1, 2).setValue(part);            // recipientsにパートを入力する
    var name      = registerSheet.getRange(i+1, 4).getValue();  // register_form_answerから氏名を取得する
    recipientsSheet.getRange(i+1, 3).setValue(name);            // recipientsに氏名を入力する
    var mailto    = registerSheet.getRange(i+1, 5).getValue();  // register_form_answerからメールアドレスを取得する
    recipientsSheet.getRange(i+1, 4).setValue(mailto);          // recipientsにメールアドレスを入力する
  }
}

/*対象者にフォームのリンクを付したメールを一斉送信する*/
function sendEveryoneEveryday () {
  var recipientsSheet   = SpreadsheetApp.openById(recipientsSheetId).getActiveSheet();
  var nRecipients       = recipientsSheet.getLastRow()-1;
  var todayCode         = Utilities.formatDate(new Date(), 'GMT+9', 'yyyyMMdd')           // 日付コード
  var todayYear         = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy')               // 現在年
  var todayMonth        = Utilities.formatDate(new Date(), 'GMT+9', 'MM')                 // 現在月
  var todayDate         = Utilities.formatDate(new Date(), 'GMT+9', 'dd')                 // 現在日
  var todayMd           = Utilities.formatDate(new Date(), 'GMT+9', 'MM月dd日')            // 日付（日本語）

  /*各参加者に下記内容のメールを送信する*/
  /*'\n'で改行、'\n\n'で2行分改行（1行空け）*/
  for (var i=1; i<=nRecipients; i++) {
    var mailto    = recipientsSheet.getRange(i+1, 4).getValue();
    var subject   = '【健康観察システム】'+ todayMd +'の報告';
    var studentno = recipientsSheet.getRange(i+1, 1).getValue();
    var part      = recipientsSheet.getRange(i+1, 2).getValue();
    var name      = recipientsSheet.getRange(i+1, 3).getValue();
    var body
                  =  '※このメールはシステムからの自動送信です。原則として返信なさらぬようお願いいたします。\n\n'
                  + name+'様\n\n'
                  + 'ごきげんよう。\n'
                  + '下記リンク先のあなた専用フォームから、'+ todayMd +'の健康観察を報告してください。\n\n'
                  + 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+mailto+'&entry.xxxxxxxxxx_year='+todayYear+'&entry.xxxxxxxxxx_month='+todayMonth+'&entry.xxxxxxxxxx_day='+todayDate+'\n\n'
                  // entry.xxxxxxxxxxの値は作成した報告フォームのソースコード（HTML）を参照して、対応する値をそれぞれ設定する。
                  + '毎日欠かさず報告してください。なお、システムにおける自動集計処理上の締切は特に設けておりません。\n'
                  + 'もし当日中に報告が間に合わなかった場合は、メール配信されたフォームから適宜追報告するようにしてください。自動的に集計に追加されます。\n\n'
                  + 'ただし、お身体の具合があまりよろしくない場合や、ご自身や同居されている方などに感染が疑われる場合（検査結果『陽性』など）は、団体全体の活動にかかわりますので可及的速やかに報告してください。\n\n'
                  + 'お忙しいところ大変恐れ入りますが、よろしくお願いいたします。\n'
                  + 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
                  + '===============================\n'
                  + ' 健康観察システム\n'
                  + ' health-report@example.com\n'
                  + '===============================\n\n'
                  + 'Report ID: '+todayCode+'_'+studentno;

    GmailApp.sendEmail(mailto, subject, body,
        {
            from: 'health-report@example.com', //Gmailから送信できるメールアドレスを設定する。
            name: '健康観察システム'
        });
  }
}
