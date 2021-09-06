/*当日未報告者判定用ワークシートを作成（健康観察開始前）*/
function createUnsubmittedWorksheet() {
    Logger.log('createUnsubmittedWorksheet() debug start');
    recipientsWorksheet.copyTo(registerSheet).setName('unsubmitted');
  }
  
  /*当日未報告者判定用ワークシートを更新（健康観察開始後）*/
  function renewUnsubmittedWorksheet() {
    Logger.log('renewUnsubmittedWorksheet() debug start');
    const unsubmittedWorksheet = registerSheet.getSheetByName('unsubmitted');
  
    unsubmittedWorksheet.clear();
    unsubmittedWorksheet.getRange(1, 1, 1, 1).setValue('学籍番号')
    unsubmittedWorksheet.getRange(1, 2, 1, 1).setValue('パート')
    unsubmittedWorksheet.getRange(1, 3, 1, 1).setValue('氏名')
    unsubmittedWorksheet.getRange(1, 4, 1, 1).setValue('メールアドレス')
    unsubmittedWorksheet.getRange(1, 5, 1, 1).setValue('携帯電話番号')
  
    const nRecipients = recipientsWorksheet.getLastRow()-1;
  
    for (var i=1; i<=nRecipients; i++){
      var studentno   = recipientsWorksheet.getRange(i+1, 1).getValue();
      unsubmittedWorksheet.getRange(i+1, 1).setValue(studentno);
      var part        = recipientsWorksheet.getRange(i+1, 2).getValue();
      unsubmittedWorksheet.getRange(i+1, 2).setValue(part);
      var name        = recipientsWorksheet.getRange(i+1, 3).getValue();
      unsubmittedWorksheet.getRange(i+1, 3).setValue(name);
      var email       = recipientsWorksheet.getRange(i+1, 4).getValue();
      unsubmittedWorksheet.getRange(i+1, 4).setValue(email);
      var mobilephone = recipientsWorksheet.getRange(i+1, 5).getValue();
      unsubmittedWorksheet.getRange(i+1, 5).setNumberFormat('@').setValue(mobilephone);
    }
  
    const recipientsToday = unsubmittedWorksheet.getRange(2, 1, nRecipients, 5);
    recipientsToday.sort([
      {column: 2, ascending: true},
      {column: 1, ascending: true}
    ]);
  }
