/*提出状況管理ワークシートの作成*/
function createSubmissionStatusWorksheet() {
    Logger.log('createSubmissionStatusWorksheet() debug start')
  
    unsubmittedWorksheet.copyTo(reportSheet).setName('submission_status');
    const submissionStatusWorksheet = reportSheet.getSheetByName('submission_status');
  
    submissionStatusWorksheet.deleteColumns(4, 2)
    submissionStatusWorksheet.insertColumnsAfter(3, 16)
    submissionStatusWorksheet.getRange(1, 4, 1, 1).setValue('=TODAY()-14')
    submissionStatusWorksheet.getRange(1, 5, 1, 1).setValue('=TODAY()-13')
    submissionStatusWorksheet.getRange(1, 6, 1, 1).setValue('=TODAY()-12')
    submissionStatusWorksheet.getRange(1, 7, 1, 1).setValue('=TODAY()-11')
    submissionStatusWorksheet.getRange(1, 8, 1, 1).setValue('=TODAY()-10')
    submissionStatusWorksheet.getRange(1, 9, 1, 1).setValue('=TODAY()-9')
    submissionStatusWorksheet.getRange(1, 10, 1, 1).setValue('=TODAY()-8')
    submissionStatusWorksheet.getRange(1, 11, 1, 1).setValue('=TODAY()-7')
    submissionStatusWorksheet.getRange(1, 12, 1, 1).setValue('=TODAY()-6')
    submissionStatusWorksheet.getRange(1, 13, 1, 1).setValue('=TODAY()-5')
    submissionStatusWorksheet.getRange(1, 14, 1, 1).setValue('=TODAY()-4')
    submissionStatusWorksheet.getRange(1, 15, 1, 1).setValue('=TODAY()-3')
    submissionStatusWorksheet.getRange(1, 16, 1, 1).setValue('=TODAY()-2')
    submissionStatusWorksheet.getRange(1, 17, 1, 1).setValue('=TODAY()-1')
    submissionStatusWorksheet.getRange(1, 18, 1, 1).setValue('=TODAY()')
    submissionStatusWorksheet.getRange(1, 19, 1, 1).setValue('3日間提出率')
    submissionStatusWorksheet.getRange(1, 20, 1, 1).setValue('14日間提出率')
  
    const nRecipients = submissionStatusWorksheet.getLastRow()-1;
  
    for (var i=1; i<=nRecipients; i++){
      var submission14DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=D$1)=$A'+(i+1)+',"OK")';
      var submission13DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=E$1)=$A'+(i+1)+',"OK")';
      var submission12DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=F$1)=$A'+(i+1)+',"OK")';
      var submission11DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=G$1)=$A'+(i+1)+',"OK")';
      var submission10DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=H$1)=$A'+(i+1)+',"OK")';
      var submission9DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=I$1)=$A'+(i+1)+',"OK")';
      var submission8DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=J$1)=$A'+(i+1)+',"OK")';
      var submission7DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=K$1)=$A'+(i+1)+',"OK")';
      var submission6DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=L$1)=$A'+(i+1)+',"OK")';
      var submission5DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=M$1)=$A'+(i+1)+',"OK")';
      var submission4DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=N$1)=$A'+(i+1)+',"OK")';
      var submission3DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=O$1)=$A'+(i+1)+',"OK")';
      var submission2DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=P$1)=$A'+(i+1)+',"OK")';
      var submissionYesterday       = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=Q$1)=$A'+(i+1)+',"OK")';
      var submissionToday           = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=R$1)=$A'+(i+1)+',"OK")';
      var last3DaysSubmissionRate   = '=COUNTIF(Q'+(i+1)+':R'+(i+1)+', "OK")/COUNTA(Q'+(i+1)+':R'+(i+1)+')';
      var last14DaysSubmissionRate  = '=COUNTIF(D'+(i+1)+':R'+(i+1)+', "OK")/COUNTA(D'+(i+1)+':R'+(i+1)+')';
  
      submissionStatusWorksheet.getRange(i+1, 4).setValue(submission14DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 5).setValue(submission13DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 6).setValue(submission12DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 7).setValue(submission11DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 8).setValue(submission10DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 9).setValue(submission9DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 10).setValue(submission8DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 11).setValue(submission7DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 12).setValue(submission6DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 13).setValue(submission5DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 14).setValue(submission4DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 15).setValue(submission3DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 16).setValue(submission2DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 17).setValue(submissionYesterday);
      submissionStatusWorksheet.getRange(i+1, 18).setValue(submissionToday);
      submissionStatusWorksheet.getRange(i+1, 19).setValue(last3DaysSubmissionRate).setNumberFormat('0.0%');
      submissionStatusWorksheet.getRange(i+1, 20).setValue(last14DaysSubmissionRate).setNumberFormat('0.0%');
  
      submissionStatusWorksheet.hideColumns(4, 12);
    }
  }
  
  /*提出状況管理ワークシートの更新*/
  function renewSubmissionStatusWorksheet() {
    Logger.log('renewSubmissionStatusWorksheet() debug start')
  
    const submissionStatusWorksheet = reportSheet.getSheetByName('submission_status');
  
    submissionStatusWorksheet.clear();
    submissionStatusWorksheet.getRange(1, 1, 1, 1).setValue('学籍番号');
    submissionStatusWorksheet.getRange(1, 2, 1, 1).setValue('パート');
    submissionStatusWorksheet.getRange(1, 3, 1, 1).setValue('氏名');
    submissionStatusWorksheet.getRange(1, 4, 1, 1).setValue('=TODAY()-14')
    submissionStatusWorksheet.getRange(1, 5, 1, 1).setValue('=TODAY()-13')
    submissionStatusWorksheet.getRange(1, 6, 1, 1).setValue('=TODAY()-12')
    submissionStatusWorksheet.getRange(1, 7, 1, 1).setValue('=TODAY()-11')
    submissionStatusWorksheet.getRange(1, 8, 1, 1).setValue('=TODAY()-10')
    submissionStatusWorksheet.getRange(1, 9, 1, 1).setValue('=TODAY()-9')
    submissionStatusWorksheet.getRange(1, 10, 1, 1).setValue('=TODAY()-8')
    submissionStatusWorksheet.getRange(1, 11, 1, 1).setValue('=TODAY()-7')
    submissionStatusWorksheet.getRange(1, 12, 1, 1).setValue('=TODAY()-6')
    submissionStatusWorksheet.getRange(1, 13, 1, 1).setValue('=TODAY()-5')
    submissionStatusWorksheet.getRange(1, 14, 1, 1).setValue('=TODAY()-4')
    submissionStatusWorksheet.getRange(1, 15, 1, 1).setValue('=TODAY()-3')
    submissionStatusWorksheet.getRange(1, 16, 1, 1).setValue('=TODAY()-2')
    submissionStatusWorksheet.getRange(1, 17, 1, 1).setValue('=TODAY()-1')
    submissionStatusWorksheet.getRange(1, 18, 1, 1).setValue('=TODAY()')
    submissionStatusWorksheet.getRange(1, 19, 1, 1).setValue('3日間提出率')
    submissionStatusWorksheet.getRange(1, 20, 1, 1).setValue('14日間提出率')
  
    const nRecipients = unsubmittedWorksheet.getLastRow()-1;
  
    for (var i=1; i<=nRecipients; i++){
      var studentno = unsubmittedWorksheet.getRange(i+1, 1).getValue();
      submissionStatusWorksheet.getRange(i+1, 1).setValue(studentno);
      var part      = unsubmittedWorksheet.getRange(i+1, 2).getValue();
      submissionStatusWorksheet.getRange(i+1, 2).setValue(part);
      var name      = unsubmittedWorksheet.getRange(i+1, 3).getValue();
      submissionStatusWorksheet.getRange(i+1, 3).setValue(name);
    }
  
    for (var i=1; i<=nRecipients; i++){
      var submission14DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=D$1)=$A'+(i+1)+',"OK")';
      var submission13DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=E$1)=$A'+(i+1)+',"OK")';
      var submission12DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=F$1)=$A'+(i+1)+',"OK")';
      var submission11DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=G$1)=$A'+(i+1)+',"OK")';
      var submission10DaysBefore    = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=H$1)=$A'+(i+1)+',"OK")';
      var submission9DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=I$1)=$A'+(i+1)+',"OK")';
      var submission8DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=J$1)=$A'+(i+1)+',"OK")';
      var submission7DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=K$1)=$A'+(i+1)+',"OK")';
      var submission6DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=L$1)=$A'+(i+1)+',"OK")';
      var submission5DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=M$1)=$A'+(i+1)+',"OK")';
      var submission4DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=N$1)=$A'+(i+1)+',"OK")';
      var submission3DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=O$1)=$A'+(i+1)+',"OK")';
      var submission2DaysBefore     = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=P$1)=$A'+(i+1)+',"OK")';
      var submissionYesterday       = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=Q$1)=$A'+(i+1)+',"OK")';
      var submissionToday           = '=IF(FILTER(report_form_answer!$B:$G, report_form_answer!$B:$B=$A'+(i+1)+', report_form_answer!$C:$C=$B'+(i+1)+', report_form_answer!$D:$D=$C'+(i+1)+', report_form_answer!$G:$G=R$1)=$A'+(i+1)+',"OK")';
      var last3DaysSubmissionRate   = '=COUNTIF(Q'+(i+1)+':R'+(i+1)+', "OK")/COUNTA(Q'+(i+1)+':R'+(i+1)+')';
      var last14DaysSubmissionRate  = '=COUNTIF(D'+(i+1)+':R'+(i+1)+', "OK")/COUNTA(D'+(i+1)+':R'+(i+1)+')';
  
      submissionStatusWorksheet.getRange(i+1, 4).setValue(submission14DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 5).setValue(submission13DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 6).setValue(submission12DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 7).setValue(submission11DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 8).setValue(submission10DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 9).setValue(submission9DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 10).setValue(submission8DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 11).setValue(submission7DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 12).setValue(submission6DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 13).setValue(submission5DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 14).setValue(submission4DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 15).setValue(submission3DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 16).setValue(submission2DaysBefore);
      submissionStatusWorksheet.getRange(i+1, 17).setValue(submissionYesterday);
      submissionStatusWorksheet.getRange(i+1, 18).setValue(submissionToday);
      submissionStatusWorksheet.getRange(i+1, 19).setValue(last3DaysSubmissionRate).setNumberFormat('0.0%');
      submissionStatusWorksheet.getRange(i+1, 20).setValue(last14DaysSubmissionRate).setNumberFormat('0.0%');
  
      submissionStatusWorksheet.hideColumns(4, 12);
  
    }
  }
  