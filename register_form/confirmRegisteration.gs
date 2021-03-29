/*ユーザに登録確認メールを送信する*/
// FormApp.getActiveForm()
function sendRegisterConfirm(e){
    Logger.log('sendRegisterConfirm() debug start');

    const responses         = e.response.getItemResponses();
    const studentno         = responses[0].getResponse();
    const part              = responses[1].getResponse();
    const name              = responses[2].getResponse();
    const email             = responses[3].getResponse();
    const consent           = responses[4].getResponse();

    var todayCode         = Utilities.formatDate(new Date(), 'GMT+9', 'yyyyMMdd');

    let mailTo      = email;
    let mailCc      = '';
    let mailBcc     = '';
    let mailReplyTo = '';
    const subject = '【健康観察システム】登録完了通知'
    var body
    = '※このメールはシステムからの自動送信です。原則として返信なさらぬようお願いいたします。\n\n'
    + name+'様\n\n'
    + '登録をいただきありがとうございます。\n'
    + '以下の内容でシステムへの登録を承りました。\n\n'
    +'------------------------------------------------------------\n\n';
      
      for (var i = 0; i < responses.length; i++) {
        var item = responses[i];
        var q = item.getItem().getTitle();
        var a = item.getResponse();
        body += '【'+ q + '】' + '\n' + '　' + a + '\n\n'
        }

    const footer
    = '------------------------------------------------------------\n\n'
    + '健康観察対象期間になりましたら、こちらのメールアドレスにあなた専用の健康観察フォームのURLをお送りします。\n\n'
    + 'お忙しいところ恐れ入りますが、ご協力のほどよろしくお願いいたします。\n'
    + 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
    + '===============================\n'
    + ' 健康観察システム\n'
    + ' health-report@example.com\n'
    + '===============================\n\n'
    + 'Register ID: '+todayCode+'_'+studentno+'_'+part+'_'+name;
    
    GmailApp.sendEmail(mailTo, subject, body+footer,
        {
          cc: mailCc,
          bcc: mailBcc, 
          from: 'health-report@example.com', //Gmailから送信できるメールアドレスを設定する。
          name: '健康観察システム',
          replyTo: mailReplyTo
        });
}
