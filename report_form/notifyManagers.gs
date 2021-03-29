/*体調不良者を役職者に通知する*/
// FormApp.getActiveForm()
function notifyManagers(e){
    const responses         = e.response.getItemResponses();
    const studentno         = responses[0].getResponse();
    const part              = responses[1].getResponse();
    const name              = responses[2].getResponse();
    const email             = responses[3].getResponse();
    const reportdate        = responses[4].getResponse();
    const temperature       = responses[5].getResponse();
    const condition         = responses[6].getResponse();
    //const conditioncategory = responses[7].getResponse();
    //const conditiondetails  = responses[8].getResponse();
    //const attendance        = responses[9].getResponse();
    //const notes             = responses[10].getResponse();
  
    var todayCode         = Utilities.formatDate(new Date(), 'GMT+9', 'yyyyMMdd');
      
    // 通報メールの設定
    let mailTo      = 'health-report@example.com';
    let mailCc      = 'manager1@example.com,manager2@example.com';
    let mailBcc     = '';
    let mailReplyTo = email;
    const subject = '【至急 健康観察システム】体調不良等発生'
    if (temperature>=37.5 || condition==='あり') {
      var body
      = '※このメールは37.5℃を超える発熱または体調不良等が報告されたことによるシステムからの自動送信です。\n\n'
      + 'cc: 管理者1、管理者2\n'
      + 'replyto: '+studentno+' '+part+' '+name+'（'+email+'）\n\n'
      + '健康観察システムにおいて、体調不良等が報告されました。\n'
      + '内容を確認のうえ、至急対応してください。\n'
      + '詳細は以下の通りです。\n'
      + '------------------------------------------------------------\n\n';
      
      for (var i = 0; i < responses.length; i++) {
        var item = responses[i];
        var q = item.getItem().getTitle();
        var a = item.getResponse();
        body += '【'+ q + '】' + '\n' + '　' + a + '\n\n'
        }
      
      const footer
      = '------------------------------------------------------------\n\n'
      + 'なお、このメールに返信すると体調不良等を報告したメンバーに対して直接連絡をとることが可能です。\n'
      + 'また、CCに含まれている関係者間での情報共有のため、可能であれば「全員に返信」する形で送信してください。\n\n'
      + 'ご不明な点などございましたら、下記メールアドレスまでお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
      + 'Report ID: '+todayCode+'_'+studentno;
    
      GmailApp.sendEmail(mailTo, subject, body+footer,
        {
          cc: mailCc,
          bcc: mailBcc,
          from: 'health-report@example.com',
          name: '健康観察システム',
          replyTo: mailReplyTo
        });
      } else {
    }
}
