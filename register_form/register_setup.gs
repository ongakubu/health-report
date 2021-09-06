/*変数定義と設定*/
/*時間関係コードの定義*/
const today         = new Date();
const todayYear     = Utilities.formatDate(today, 'GMT+9', 'yyyy');
const todayMonth    = Utilities.formatDate(today, 'GMT+9', 'MM');
const todayDate     = Utilities.formatDate(today, 'GMT+9', 'dd');
const todayMd       = Utilities.formatDate(today, 'GMT+9', 'M月d日');
const todayCode     = Utilities.formatDate(today, 'GMT+9', 'yyyyMMdd');
const currentTime   = Utilities.formatDate(today, 'GMT+9', 'H:mm');
const yesterday     = new Date(today.getFullYear(), today.getMonth(), today.getDate()-1);
const yesterdayCode = Utilities.formatDate(yesterday, 'GMT+9', 'yyyyMMdd');

const startDate = new Date();
/*setFullYear(西暦4桁, 月0–11, 日1–31)*/
startDate.setFullYear(2021, 9, 31);  /*例えばこれは2021年10月31日*/
startDate.setHours(4);
startDate.setMinutes(0);
startDate.setSeconds(0);

const maintenanceTime = new Date();
maintenanceTime.setHours(5);
maintenanceTime.setMinutes(55);

const deadline  = new Date();
deadline.setFullYear(2021, 10, 20);
deadline.setHours(9);
deadline.setMinutes(0);
deadline.setSeconds(0);
const nMunitutesBeforeForReminder1 = 30;
const nMunitutesBeforeForReminder2 = 15;

/*各ファイルの定義*/
const registerFormId    = '登録フォームのID';
const registerForm      = FormApp.openById(registerFormId);
const registerSheetId   = '登録フォーム（回答）スプレッドシートのID';
const registerSheet     = SpreadsheetApp.openById(registerSheetId);

const reportFormId    = '報告フォームのID';
const reportForm      = FormApp.openById(reportFormId);
const reportSheetId   = '報告フォーム（回答）スプレッドシートのID';
const reportSheet     = SpreadsheetApp.openById(reportSheetId);

/*各メールアドレスの定義*/
const systemEmail       = 'health-report@example.com';
const supervisorsEmail  = 'manager1@example.com,manager2@example.com';

function initialization() {
  Logger.log('initialization() debug start');
  registerSheet.getSheets()[0].setName('register_form_answer');
  reportSheet.getSheets()[0].setName('report_form_answer');

  let projectTriggers =ScriptApp.getProjectTriggers();
  for (const trigger of projectTriggers) {
    if (trigger.getHandlerFunction() != null){
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger('manageRegistration').forForm(registerForm).onFormSubmit().create();

  const triggerTime = new Date();
  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate()-15);
  triggerTime.setHours(0);
  triggerTime.setMinutes(0);
  ScriptApp.newTrigger('disableDeregisterOption').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate()+15);
  triggerTime.setHours(0);
  triggerTime.setMinutes(0);
  ScriptApp.newTrigger('enableDeregisterOption').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  triggerTime.setHours(deadline.getHours());
  triggerTime.setMinutes(deadline.getMinutes()-nMunitutesBeforeForReminder1);
  ScriptApp.newTrigger('sendBulkReminder1').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  triggerTime.setHours(deadline.getHours());
  triggerTime.setMinutes(deadline.getMinutes()-nMunitutesBeforeForReminder2);
  ScriptApp.newTrigger('sendBulkReminder2').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  triggerTime.setHours(deadline.getHours());
  triggerTime.setMinutes(deadline.getMinutes()+1);
  ScriptApp.newTrigger('closeFormAcceptanceDeadline').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  triggerTime.setHours(deadline.getHours());
  triggerTime.setMinutes(deadline.getMinutes()+5);
  ScriptApp.newTrigger('openFormAcceptanceDeadline').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  triggerTime.setHours(deadline.getHours());
  triggerTime.setMinutes(deadline.getMinutes()+5);
  ScriptApp.newTrigger('sendBulkUrgentReminder').timeBased().at(triggerTime).create();

  registerSheet.insertSheet().setName('recipients');
  const recipientsWorksheet = registerSheet.getSheetByName('recipients');
  recipientsWorksheet.getRange(1, 1, 1, 1).setValue('学籍番号');
  recipientsWorksheet.getRange(1, 2, 1, 1).setValue('パート');
  recipientsWorksheet.getRange(1, 3, 1, 1).setValue('氏名');
  recipientsWorksheet.getRange(1, 4, 1, 1).setValue('メールアドレス');
  recipientsWorksheet.getRange(1, 5, 1, 1).setValue('携帯電話番号');
  recipientsWorksheet.deleteColumns(6, 21);
  recipientsWorksheet.deleteRows(2, 998);
  
  triggerTime.setFullYear(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  triggerTime.setHours(startDate.getHours());
  triggerTime.setMinutes(startDate.getMinutes());
  ScriptApp.newTrigger('createUnsubmittedWorksheet').timeBased().at(triggerTime).create();

  ScriptApp.newTrigger('setDailyTrigger').timeBased().atHour(maintenanceTime.getHours()-1).everyDays(1).create();
}

function setDailyTrigger() {
  Logger.log('setDailyMaintenanceTrigger() debug start');
  let projectTriggers =ScriptApp.getProjectTriggers();
  for (const trigger of projectTriggers) {
    if (trigger.getHandlerFunction() == 'closeFormAcceptance' || trigger.getHandlerFunction() == 'renewUnsubmittedWorksheet' || trigger.getHandlerFunction() == 'openFormAcceptance' || trigger.getHandlerFunction() == 'sendBulk'){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const triggerTime = new Date();
  triggerTime.setHours(maintenanceTime.getHours());
  triggerTime.setMinutes(maintenanceTime.getMinutes()-15);
  ScriptApp.newTrigger('closeFormAcceptance').timeBased().at(triggerTime).create();

  triggerTime.setHours(maintenanceTime.getHours());
  triggerTime.setMinutes(maintenanceTime.getMinutes()-13);
  ScriptApp.newTrigger('renewUnsubmittedWorksheet').timeBased().at(triggerTime).create();

  triggerTime.setHours(maintenanceTime.getHours());
  triggerTime.setMinutes(maintenanceTime.getMinutes());
  ScriptApp.newTrigger('openFormAcceptance').timeBased().at(triggerTime).create();

  triggerTime.setHours(maintenanceTime.getHours());
  triggerTime.setMinutes(maintenanceTime.getMinutes()+5);
  ScriptApp.newTrigger('sendBulk').timeBased().at(triggerTime).create();
}

/*配信宛先管理などのワークシート操作の途中で登録や報告の送信をさせないように、フォームを一時受付停止*/
function closeFormAcceptance() {
  Logger.log('closeFormAcceptance() debug start');
  const maintenanceMessage = '定時メンテナンス中です。予定通りにいけば'+maintenanceTime.getHours()+'時'+maintenanceTime.getMinutes()+'分頃には完了する見込みです。\n大変恐れ入りますが、しばらくしてからもう一度アクセスしてください。\n\nお急ぎの場合は下記メールアドレスまでご連絡ください。\n\n＜お問い合わせ先＞\n健康観察システム担当\nhealth-report@example.com';
  registerForm.setAcceptingResponses(false).setCustomClosedFormMessage(maintenanceMessage);
  reportForm.setAcceptingResponses(false).setCustomClosedFormMessage(maintenanceMessage);
}
/*フォームの一時受付停止を解除し、受付を再開*/
function openFormAcceptance() {
  Logger.log('openFormAcceptance() debug start');
  registerForm.setAcceptingResponses(true);
  reportForm.setAcceptingResponses(true);
}

/*演奏会当日の締切時刻周辺でフォームの一時受付停止*/
function closeFormAcceptanceDeadline() {
  Logger.log('closeFormAcceptanceDeadline() debug start');
  const closedMessage = deadline.getHours()+'時'+deadline.getMinutes()+'分を過ぎましたので、健康観察報告の受付を締め切りました。\n\nご不明な点などございましたら、下記メールアドレスまでお気軽にお問い合わせください。\n\n＜お問い合わせ先＞\n健康観察システム担当\nhealth-report@example.com';
  registerForm.setAcceptingResponses(false).setCustomClosedFormMessage(closedMessage);
  reportForm.setAcceptingResponses(false).setCustomClosedFormMessage(closedMessage);
}
/*フォームの一時受付停止を解除し、受付を再開*/
function openFormAcceptanceDeadline() {
  Logger.log('openFormAcceptanceDeadline() debug start');
  registerForm.setAcceptingResponses(true);
  reportForm.setAcceptingResponses(true);
}

/*システムからの解除の受付を無効化*/
function disableDeregisterOption() {
  Logger.log('disableDeregisterOption() debug start');
  const registerFormQuestions = registerForm.getItems();
  const overwriteQuestion = registerFormQuestions[7];
  const overwriteOptions = [];
    overwriteOptions[0] = '新規';
    overwriteOptions[1] = '更新';
    overwriteOptions[2] = '照会';
  overwriteQuestion.asListItem().setChoiceValues(overwriteOptions);
}
/*システムからの解除の受付を有効化*/
function enableDeregisterOption() {
  Logger.log('enableDeregisterOption() debug start');
  const registerFormQuestions = registerForm.getItems();
  const overwriteQuestion = registerFormQuestions[7];
  const overwriteOptions = [];
    overwriteOptions[0] = '新規';
    overwriteOptions[1] = '更新';
    overwriteOptions[2] = '解除';
    overwriteOptions[3] = '照会';
  overwriteQuestion.asListItem().setChoiceValues(overwriteOptions);
}
