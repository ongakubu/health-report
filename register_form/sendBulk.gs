const unsubmittedWorksheet = registerSheet.getSheetByName('unsubmitted');

/*対象者に報告フォームのリンクを付したメールを一斉配信*/
function sendBulk() {
  Logger.log('sendBulk() debug start')

  const nRecipients = unsubmittedWorksheet.getLastRow()-1;

  /*配信メールの文面作成*/
  for (var i=1; i<=nRecipients; i++) {
    var mailTo      = unsubmittedWorksheet.getRange(i+1, 4).getValue();
    let mailCc      = '';
    let mailBcc     = '';
    let mailReplyTo = '';

    var studentno   = unsubmittedWorksheet.getRange(i+1, 1).getValue();
    var part        = unsubmittedWorksheet.getRange(i+1, 2).getValue();
    var name        = unsubmittedWorksheet.getRange(i+1, 3).getValue();
    var mobilephone = unsubmittedWorksheet.getRange(i+1, 5).getValue();

    var subject     = '【健康観察】'+ todayMd +'の報告';
    var body
      = '※このメールはシステムからの自動送信です。原則として返信なさらぬようお願いいたします。\n\n'
      + name+'様\n\n'
      + 'ごきげんよう。\n'
      + '下記リンク先のあなた専用フォームから、'+ todayMd +'の健康観察を報告してください。\n\n'
      + 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+mailTo+'&entry.xxxxxxxxx='+mobilephone+'&entry.xxxxxxxxx='+todayYear+'-'+todayMonth+'-'+todayDate+'\n\n'
      + '必ず毎日報告してください。\n'
      + 'もし当日中に報告が間に合わなかった場合は、メール配信されたフォームから適宜追報告を行ってください。\n\n'
      + 'また、お身体の具合があまりよろしくない場合や、ご自身や同居されている方などに感染が疑われる場合（検査結果『陽性』など）は、あなた自身を含む全ての団体構成員の健康と活動にかかわりますので可及的速やかに報告してください。\n\n'
      + 'お忙しいところ大変恐れ入りますが、よろしくお願いいたします。\n\n';

    const footer
      = 'なお、システムに登録されているメールアドレスなどを更新したり、今後の配信を希望しないため登録を解除したり、これまで報告した健康観察の記録を照会したい場合は、下記リンク先の登録フォームからお申し込みください。\n\n'
      + 'https://docs.google.com/forms/d/e/登録フォームのID/viewform?entry.xxxxxxxxxx='+studentno+'&entry.xxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'\n\n'
      + 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
      + 'Mail ID: '+todayCode+'_'+studentno+'_'+part+'_'+name;

  /*配信メールの送信*/
  GmailApp.sendEmail(mailTo, subject, body+footer,
    {
      cc: mailCc,
      bcc: mailBcc, 
      from: systemEmail,
      name: '健康観察システム',
      replyTo: mailReplyTo
    });
  }
}

/*未報告者に報告フォームのリンクを付したリマインダメールを一斉配信*/
function sendBulkReminder1() {
  Logger.log('sendBulkReminder1() debug start')

  const nRecipients = unsubmittedWorksheet.getLastRow()-1;

  /*リマインダメールの文面作成*/
  for (var i=1; i<=nRecipients; i++) {
    var mailTo      = unsubmittedWorksheet.getRange(i+1, 4).getValue();
    let mailCc      = '';
    let mailBcc     = '';
    let mailReplyTo = '';

    var studentno   = unsubmittedWorksheet.getRange(i+1, 1).getValue();
    var part        = unsubmittedWorksheet.getRange(i+1, 2).getValue();
    var name        = unsubmittedWorksheet.getRange(i+1, 3).getValue();
    var mobilephone = unsubmittedWorksheet.getRange(i+1, 5).getValue();

    var subject     = '【健康観察】あと'+nMunitutesBeforeForReminder1+'分ほどで締切です';
    var body
      = '※このメールは'+currentTime+'時点で本日分の健康観察が未報告の方に向けてシステムが自動で送信したものです（タイミング的に行き違いになってしまった方は大変申し訳ございません）。\n\n'
      + name+'様\n\n'
      + 'ごきげんよう。\n'
      + '必ず'+deadline.getHours()+'時'+deadline.getMinutes()+'分までに下記リンク先のあなた専用フォームから、'+ todayMd +'の健康観察を報告してください。\n\n'
      + 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+mailTo+'&entry.xxxxxxxxx='+mobilephone+'&entry.xxxxxxxxx='+todayYear+'-'+todayMonth+'-'+todayDate+'\n\n'
      + 'あと'+nMunitutesBeforeForReminder1+'分ほどで受付が自動で締め切られ、それ以降の回答を一切受け付けません。\n\n'
      + 'お忙しいところ大変恐れ入りますが、どうかよろしくお願いいたします。\n\n';
    const footer
      = 'なお、システムに登録されているメールアドレスなどを更新したり、今後の配信を希望しないため登録を解除したり、これまで報告した健康観察の記録を照会したい場合は、下記リンク先の登録フォームからお申し込みください。\n\n'
      + 'https://docs.google.com/forms/d/e/登録フォームのID/viewform?entry.xxxxxxxxxx='+studentno+'&entry.xxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'\n\n'
      + 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
      + 'Mail ID: '+todayCode+'_'+studentno+'_'+part+'_'+name+'_Reminder1';

  /*リマインダメールの送信*/
  GmailApp.sendEmail(mailTo, subject, body+footer,
    {
      cc: mailCc,
      bcc: mailBcc, 
      from: systemEmail,
      name: '健康観察システム',
      replyTo: mailReplyTo
    });
  }
}

/*未報告者に報告フォームのリンクを付したリマインダメールを一斉配信*/
function sendBulkReminder2() {
  Logger.log('sendBulkReminder2() debug start')

  const nRecipients = unsubmittedWorksheet.getLastRow()-1;

  /*リマインダメールの文面作成*/
  for (var i=1; i<=nRecipients; i++) {
    var mailTo      = unsubmittedWorksheet.getRange(i+1, 4).getValue();
    let mailCc      = '';
    let mailBcc     = '';
    let mailReplyTo = '';

    var studentno   = unsubmittedWorksheet.getRange(i+1, 1).getValue();
    var part        = unsubmittedWorksheet.getRange(i+1, 2).getValue();
    var name        = unsubmittedWorksheet.getRange(i+1, 3).getValue();
    var mobilephone = unsubmittedWorksheet.getRange(i+1, 5).getValue();

    var subject     = '【健康観察】あと'+nMunitutesBeforeForReminder2+'分ほどで締切です';
    var body
      = '※このメールは'+currentTime+'時点で本日分の健康観察が未報告の方に向けてシステムが自動で送信したものです（タイミング的に行き違いになってしまった方は大変申し訳ございません）。\n\n'
      + name+'様\n\n'
      + 'ごきげんよう。\n'
      + '必ず'+deadline.getHours()+'時'+deadline.getMinutes()+'分までに下記リンク先のあなた専用フォームから、'+ todayMd +'の健康観察を報告してください。\n\n'
      + 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+mailTo+'&entry.xxxxxxxxx='+mobilephone+'&entry.xxxxxxxxx='+todayYear+'-'+todayMonth+'-'+todayDate+'\n\n'
      + 'あと'+nMunitutesBeforeForReminder2+'分ほどで受付が自動で締め切られ、それ以降の回答を一切受け付けません。\n\n'
      + 'お忙しいところ大変恐れ入りますが、どうかよろしくお願いいたします。\n\n';
    const footer
      = 'なお、システムに登録されているメールアドレスなどを更新したり、今後の配信を希望しないため登録を解除したり、これまで報告した健康観察の記録を照会したい場合は、下記リンク先の登録フォームからお申し込みください。\n\n'
      + 'https://docs.google.com/forms/d/e/登録フォームのID/viewform?entry.xxxxxxxxxx='+studentno+'&entry.xxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'\n\n'
      + 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
      + 'Mail ID: '+todayCode+'_'+studentno+'_'+part+'_'+name+'_Reminder2';

  /*リマインダメールの送信*/
  GmailApp.sendEmail(mailTo, subject, body+footer,
    {
      cc: mailCc,
      bcc: mailBcc, 
      from: systemEmail,
      name: '健康観察システム',
      replyTo: mailReplyTo
    });
  }
}

/*締切を過ぎても提出しない未報告者に報告フォームのリンクを付した（怒りの）督促メールを一斉配信*/
function sendBulkUrgentReminder() {
  Logger.log('sendBulkUrgentReminder() debug start')

  const nRecipients = unsubmittedWorksheet.getLastRow()-1;

  /*リマインダメールの文面作成*/
  for (var i=1; i<=nRecipients; i++) {
    var mailTo      = unsubmittedWorksheet.getRange(i+1, 4).getValue();
    let mailCc      = '';
    let mailBcc     = '';
    let mailReplyTo = '';

    var studentno   = unsubmittedWorksheet.getRange(i+1, 1).getValue();
    var part        = unsubmittedWorksheet.getRange(i+1, 2).getValue();
    var name        = unsubmittedWorksheet.getRange(i+1, 3).getValue();
    var mobilephone = unsubmittedWorksheet.getRange(i+1, 5).getValue();

    var subject     = '【健康観察】'+name+'さん、絶対に報告してくださいね。';
    var body
      = '※このメールは、'+deadline.getHours()+'時'+deadline.getMinutes()+'分の締切を過ぎてもなお本日分の健康観察が未報告の'+name+'様に向けてシステムが自動で送信したものです（タイミング的に行き違いになってしまった方は大変申し訳ございません）。\n\n'
      + name+'様\n\n'
      + 'ごきげんよう。\n'
      + '必ず'+deadline.getHours()+'時'+deadline.getMinutes()+'分までに下記リンク先のあなた専用フォームから、'+ todayMd +'の健康観察を報告してくださいと3回もお願いしたのに...\n\n'
      + 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+mailTo+'&entry.xxxxxxxxx='+mobilephone+'&entry.xxxxxxxxx='+todayYear+'-'+todayMonth+'-'+todayDate+'\n\n'
      + name+'様のため報告フォームの受付を一時的に再開しましたので、絶対に報告してください。\n\n'
    const footer
      = 'なお、システムに登録されているメールアドレスなどを更新したり、今後の配信を希望しないため登録を解除したり、これまで報告した健康観察の記録を照会したい場合は、下記リンク先の登録フォームからお申し込みください。\n\n'
      + 'https://docs.google.com/forms/d/e/登録フォームのID/viewform?entry.xxxxxxxxxx='+studentno+'&entry.xxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'\n\n'
      + 'ご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n'
      + '===============================\n'
      + ' 健康観察システム\n'
      + ' health-report@example.com\n'
      + '===============================\n\n'
      + 'Mail ID: '+todayCode+'_'+studentno+'_'+part+'_'+name+'_UrgentReminder';

  /*リマインダメールの送信*/
  GmailApp.sendEmail(mailTo, subject, body+footer,
    {
      cc: mailCc,
      bcc: mailBcc, 
      from: systemEmail,
      name: '健康観察システム',
      replyTo: mailReplyTo
    });
  }
}
