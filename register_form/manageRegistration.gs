const registerWorksheet   = registerSheet.getSheetByName('register_form_answer');
const reportWorksheet     = reportSheet.getSheetByName('report_form_answer');
const recipientsWorksheet = registerSheet.getSheetByName('recipients');

/*登録フォームの回答に応じた登録完了通知メールの作成送信及びrecipientsWorksheetの書換*/
// FormApp.getActiveForm()
function manageRegistration(e){
  Logger.log('manageRegistration(e) debug start');

  /*回答内容の取得と回答項目変数の定義*/
  const responses   = e.response.getItemResponses();
  const studentno   = responses[0].getResponse();
  const part        = responses[1].getResponse();
  const name        = responses[2].getResponse();
  const email       = responses[3].getResponse();
  const mobilephone = responses[4].getResponse();
  const consent     = responses[5].getResponse();
  const overwrite   = responses[6].getResponse();

  /*登録完了通知メールの文面作成*/
  const mailTo      = email;
  const mailCc      = '';
  const mailBcc     = '';
  const mailReplyTo = '';
  const subject     = '【健康観察】登録完了通知（'+overwrite+'）';
  let body
    = '※このメールはシステムからの自動送信です。原則として返信なさらぬようお願いいたします。\n\n'
    + name+'様\n\n'
    + '登録をいただきありがとうございます。\n'
    + '以下の内容でシステムへの登録（'+overwrite+'）を承りました。\n\n'
    +'------------------------------------------------------------\n\n';
      
      for (var i = 0; i < responses.length; i++) {
        var item      = responses[i];
        var question  = item.getItem().getTitle();
        var answer    = item.getResponse();
        body += '【'+ question + '】' + '\n' + '　' + answer + '\n\n'
      }

    if (overwrite == '新規') {
      description
      = '------------------------------------------------------------\n\n'
      + '明朝から、こちらのメールアドレスにあなた専用の健康観察フォームのURLをお送りします。\n\n'
      + 'お急ぎの場合は、下記リンク先のあなた専用フォームから必要な報告日の健康観察を適宜送信してください。\n\n'
      + 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+email+'&entry.xxxxxxxxx='+mobilephone+'\n\n'
      + 'お忙しいところ恐れ入りますが、ご協力のほどよろしくお願いいたします。\n\n';
    } else if (overwrite == '更新') {
      description
      = '------------------------------------------------------------\n\n'
      + '健康観察システムに登録されている'+name+'様の情報を更新いたしました。\n'
      + '次回の健康観察報告の配信時から、更新された情報が適用されます。\n\n';
    } else if (overwrite == '解除') {
      description
      = '------------------------------------------------------------\n\n'
      + '健康観察システムから'+name+'様の登録を解除いたしました。\n'
      + 'ご利用ありがとうございました、お元気で。\n\n';
    } else if (overwrite == '照会') {
      /*部員の報告データの取得*/
      const reportData = reportWorksheet.getDataRange().getDisplayValues();
      const reportDataHeader = reportData.shift(), cIndex = {};
      for (var i = 0; i < reportDataHeader.length; i++) cIndex[reportDataHeader[i]] = i;
      const recipientsAttendanceData = reportData
        .filter(function(e){return e /*[cIndex['学籍番号']] === studentno && [cIndex['パート']] === part &&*/ [cIndex['氏名']] === name })
        .map(function(e){
          const reportDataColumns = [
            cIndex['報告日'],
            cIndex['体温'],
            cIndex['体調不良等の有無'],
            cIndex['体調不良等種別'],
          ], reportDataRow = [];
          for (var i = 0; i < reportDataColumns.length; i++) reportDataRow.push(e[reportDataColumns[i]]);
          return reportDataRow;
        })

      let tempRecipientsReportData = [];
      for(var i = 0; i < recipientsAttendanceData.length; i++) {
        tempRecipientsReportData.push(recipientsAttendanceData[i].join('\t'));
        }
        let recipientsReportDataList = tempRecipientsReportData.join('\n　');
      
      description
      = '------------------------------------------------------------\n\n'
      + '健康観察システムに報告いただいている'+name+'様のデータは次の通りです。\n\n'
      + '------------------------------------------------------------\n\n'
      + '　'+recipientsReportDataList+'\n\n'
      + '------------------------------------------------------------\n\n';
    }

  const footer
    = 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
    + '===============================\n'
    + ' 健康観察システム\n'
    + ' health-report@example.com\n'
    + '===============================\n\n'
    + 'Register ID: '+todayCode+'_'+studentno+'_'+part+'_'+name;

  /*登録完了通知メールの送信*/
  GmailApp.sendEmail(mailTo, subject, body+description+footer,
      {
        cc: mailCc,
        bcc: mailBcc, 
        from: 'health-report@example.com',
        name: '健康観察システム',
        replyTo: mailReplyTo
      });

  /*登録内容をrecipientsWorksheetに書込*/
  /*新規または更新または解除の場合にrecipientsWorksheetから既存登録を削除（重複の防止）*/
  if (overwrite != '照会') {
    const lastRow = recipientsWorksheet.getLastRow();
    for(var i = lastRow ; i > 1 ; i--){
      var cellA = recipientsWorksheet.getRange(i,1).getValue();
      var cellB = recipientsWorksheet.getRange(i,2).getValue();
      var cellC = recipientsWorksheet.getRange(i,3).getValue();
      if (cellA != studentno || cellB != part || cellC != name) {
        continue;
        } else {
          recipientsWorksheet.insertRowAfter(lastRow);
          recipientsWorksheet.deleteRows(i);
        }
      }
    } else {
  }

  /*新規または更新の場合にrecipientsWorksheetに新しい登録を追加*/
  if (overwrite == '新規' || overwrite == '更新') {
    const newRowForNewRecipients = recipientsWorksheet.getLastRow()+1;
    recipientsWorksheet.getRange(newRowForNewRecipients, 1).setValue(studentno);
    recipientsWorksheet.getRange(newRowForNewRecipients, 2).setValue(part);
    recipientsWorksheet.getRange(newRowForNewRecipients, 3).setValue(name);
    recipientsWorksheet.getRange(newRowForNewRecipients, 4).setValue(email);
    recipientsWorksheet.getRange(newRowForNewRecipients, 5).setNumberFormat('@').setValue(mobilephone);
    } else {
  }
}
