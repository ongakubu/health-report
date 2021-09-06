/*報告フォームの回答に応じた体調不良等判定ならびに体調不良等発生通報メールの作成送信及び当日未報告者判定用ワークシートの更新*/
// FormApp.getActiveForm()
function notifySupervisors(e){
    Logger.log('notifySupervisors(e) debug start')
    const responses         = e.response.getItemResponses();
    const studentno         = responses[0].getResponse();
    const part              = responses[1].getResponse();
    const name              = responses[2].getResponse();
    const email             = responses[3].getResponse();
    const mobilephone       = responses[4].getResponse();
    const reportdate        = responses[5].getResponse();
    const temperature       = responses[6].getResponse();
    const condition         = responses[7].getResponse();
    /*const conditioncategory = responses[8].getResponse();
    const conditiondetails  = responses[9].getResponse();
    const vaccination       = responses[10].getResponse();
    const vaccinationdate   = responses[11].getResponse();
    const notes             = responses[12].getResponse();*/
  
    /*体調不良等発生通報メールの文面作成*/
    /*let mailTo      = systemEmail;*/
    let mailTo      = email;
    let mailCc      = supervisorsEmail;
    let mailBcc     = systemEmail;
    let mailReplyTo = email;
    const subject = '【至急 健康観察】'+part+' '+name+'さん、お大事になさってください。'
    if (temperature >= 37.5 || condition != '無し') {
      var body
      = '※このメールは37.5°Cを超える発熱または体調不良等が報告されたことによるシステムからの自動送信です。\n\n'
      + name+'様\n\n'
      + 'cc: 管理者1、管理者2\n'
      + 'replyto: '+studentno+' '+part+' '+name+'（'+email+'）\n\n'
      + '体調不良等のご報告をいただきありがとうございます、どうかお大事になさってください。\n\n'
      + '以下の内容でシステムへの報告を承りました。\n\n'
      + '------------------------------------------------------------\n\n';
      
      for (var i = 0; i < responses.length; i++) {
          var item      = responses[i];
          var question  = item.getItem().getTitle();
          var answer    = item.getResponse();
          body += '【'+ question + '】' + '\n' + '　' + answer + '\n\n'
        }
  
      const footer
      = 'ご不明な点などございましたら、下記メールアドレスまでお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
      + 'Report ID: '+reportdate+'_'+studentno+'_'+part+'_'+name;
  
    /*体調不良等発生通報メールの送信*/
    GmailApp.sendEmail(mailTo, subject, body+description+yellowpages+footer,
      {
        cc: mailCc,
        bcc: mailBcc,
        from: systemEmail,
        name: '健康観察システム',
        replyTo: mailReplyTo
      });
      } else {
    }
  
    /*当日分を報告済の参加者をunsubmittedWorksheetから削除*/
    const unsubmittedWorksheetLastRow = unsubmittedWorksheet.getLastRow();
      for(var i = unsubmittedWorksheetLastRow ; i > 1 ; i--){
        var cellA = unsubmittedWorksheet.getRange(i,1).getValue();
        var cellB = unsubmittedWorksheet.getRange(i,2).getValue();
        var cellC = unsubmittedWorksheet.getRange(i,3).getValue();
        if (cellA != studentno || cellB != part || cellC != name || reportdate != todayYear+'-'+todayMonth+'-'+todayDate) {
          continue;
          } else {
            unsubmittedWorksheet.insertRowAfter(unsubmittedWorksheetLastRow);
            unsubmittedWorksheet.deleteRows(i);
          }
        }
}
