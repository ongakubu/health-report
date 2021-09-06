/*未報告者一覧の役職者宛通報*/
function reportUnsubmittedPerson() {
    Logger.log('reportUnsubmittedPerson() debug start')
  
    /*unsubmittedWorksheetに残っている未報告者情報を取得*/
    const nUnsubmittedPerson  = unsubmittedWorksheet.getLastRow()-1;
    if (nUnsubmittedPerson > 0) {
      for (var i=1; i<=nUnsubmittedPerson; i++) {
        const unsubmittedPerson     = unsubmittedWorksheet.getRange(2, 1, i, 3).getValues();
        var everyUnsubmittedPerson  = unsubmittedPerson.join('\n　').replace(/,/g, '\t');
      }
    } else {
        var everyUnsubmittedPerson = '全員締切までに報告しました！((o(^∇^)o))';
    };
    Logger.log(everyUnsubmittedPerson);
  
    /*未報告者通報メールの文面作成*/
    let mailTo      = systemEmail;
    let mailCc      = supervisorsEmail;
    let mailBcc     = '';
    let mailReplyTo = '';
    const subject = '【健康観察】'+currentTime+'時点の未報告者は'+nUnsubmittedPerson+'名'
    var body
      = '※このメールはシステムからの自動送信です。\n\n'
      + 'cc: 管理者1、管理者2\n\n'
      + todayMd+currentTime+'時点での未報告者は以下の通りです。\n\n'
      + '------------------------------------------------------------\n\n'
  
      + '　'+everyUnsubmittedPerson+'\n\n';
  
    const footer
      = '------------------------------------------------------------\n\n'
      + '未報告者はパートごとの学年順です（パート名がアルファベット順なのはGoogle App Scriptの仕様です）。\n\n'
      + 'ご不明な点などございましたら、下記メールアドレスまでお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
  
    /*未報告者通報メールの送信*/
    GmailApp.sendEmail(mailTo, subject, body+footer,
      {
        cc: mailCc,
        bcc: mailBcc,
        from: systemEmail,
        name: '健康観察システム',
        replyTo: mailReplyTo
      });
  }
  