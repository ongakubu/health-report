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
const registerFormId        = '登録フォームのID';
const registerForm          = FormApp.openById(registerFormId);
const registerSheetId       = '登録フォーム（回答）スプレッドシートのID';
const registerSheet         = SpreadsheetApp.openById(registerSheetId);
const registerWorksheet     = registerSheet.getSheetByName('register_form_answer');
const recipientsWorksheet   = registerSheet.getSheetByName('recipients');
const unsubmittedWorksheet  = registerSheet.getSheetByName('unsubmitted');

const reportFormId    = '報告フォームのID';
const reportForm      = FormApp.openById(reportFormId);
const reportSheetId   = '報告フォーム（回答）スプレッドシートのID';
const reportSheet     = SpreadsheetApp.openById(reportSheetId);
const reportWorksheet = reportSheet.getSheetByName('report_form_answer');

/*各メールアドレスの定義*/
const systemEmail       = 'health-report@example.com';
const supervisorsEmail  = 'manager1@example.com,manager2@example.com';

function initialization() {
  Logger.log('initialization() debug start');
  let projectTriggers =ScriptApp.getProjectTriggers();
  for (const trigger of projectTriggers) {
    if (trigger.getHandlerFunction() != null){
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger('notifySupervisors').forForm(reportForm).onFormSubmit().create();

  const triggerTime = new Date();
  triggerTime.setFullYear(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  triggerTime.setHours(startDate.getHours());
  triggerTime.setMinutes(startDate.getMinutes()+2);
  ScriptApp.newTrigger('createSubmissionStatusWorksheet').timeBased().at(triggerTime).create();

  triggerTime.setFullYear(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  triggerTime.setHours(startDate.getHours());
  triggerTime.setMinutes(startDate.getMinutes()+5);
  ScriptApp.newTrigger('createAttendanceReferenceWorksheet').timeBased().at(triggerTime).create();
  
  triggerTime.setFullYear(deadline.getFullYear(), deadline.getMonth(), deadline.getDate());
  triggerTime.setHours(deadline.getHours());
  triggerTime.setMinutes(deadline.getMinutes()+5);
  ScriptApp.newTrigger('reportUnsubmittedPerson').timeBased().at(triggerTime).create();

  ScriptApp.newTrigger('setDailyTrigger').timeBased().atHour(maintenanceTime.getHours()-1).everyDays(1).create();
}

function setDailyTrigger() {
    Logger.log('setDailyTrigger() debug start');
  let projectTriggers =ScriptApp.getProjectTriggers();
  for (const trigger of projectTriggers) {
    if (trigger.getHandlerFunction() == 'renewSubmissionStatusWorksheet'){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const triggerTime = new Date();
  triggerTime.setHours(maintenanceTime.getHours());
  triggerTime.setMinutes(maintenanceTime.getMinutes()-9);
  ScriptApp.newTrigger('renewSubmissionStatusWorksheet').timeBased().at(triggerTime).create();
}
