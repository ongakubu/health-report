# health-report
新型コロナウイルス感染症（COVID-19）対応の一環として、[学習院輔仁会音楽部](https://www.ongakubu.org)において開発・運用している健康観察システムです。
本システムを用いて、対面活動参加部員が日々行う検温や体調の報告・集計を自動化しています。
Google Forms、Spreadsheet、Apps Scriptを用いており、フォームの質問項目等は学生オーケストラでの利用を想定したものになっていますから、必要に応じて変更・調整してご利用ください。

## 主な機能
* 新規登録・登録内容更新・登録解除・報告済健康観察記録の照会手続を登録フォームから自動処理
* ユーザ識別情報を事前入力した報告フォーム（体温、体調等）をメールで毎日定時配信
* 37.5℃以上の発熱ないし体調不良等発生時に当該報告を送信したユーザと管理者へ即時通知
* 健康観察未報告者一覧を管理者へ自動通知

## 仕組み
![Structure Overview v4](https://user-images.githubusercontent.com/73869913/132485568-90b1d443-1ef8-468e-a6d1-f19cd706eefa.jpg)


1. ユーザが登録フォームに必要事項を入力して送信する。
2. 登録フォームに入力された回答を判定のうえ登録回答ワークシートに記録する。
3. 登録フォームが登録確認メールをユーザに自動送信する。
4. 登録フォームの回答をもとに宛先ワークシートを書き換える。
5. トリガーで設定した時刻に、宛先ワークシートをもとに当日未報告者一覧ワークシート（これが当日の配信対象者リストになる）を生成する。
6. トリガーで設定した時刻に、各ユーザの識別情報を事前入力した報告フォームへのリンクを含むメールを生成し、当日未報告者一覧ワークシートに記載されているメールアドレスに自動配信する。
7. ユーザが配信された報告フォームに必要事項（体温、体調等）を入力して回答を送信する。
8. 報告フォームに入力された回答を判定のうえ報告回答ワークシートに記録する
9. 当日未報告者一覧ワークシートから回答済ユーザを削除する。
10. もし発熱や体調不良等が報告された場合は、当該報告を送信したユーザと管理者にメールで自動通知する。
11. トリガーで設定した時刻に、当日未報告者一覧ワークシートの内容を管理者にメールで自動通知する（任意）。

## 構築方法
1. 登録フォームを作成する。
   ![register_form_questions.png](https://github.com/ongakubu/health-report/blob/main/screenshots/register_form_questions.png)
3. 登録フォームの回答を格納するスプレッドシートを、登録フォームの編集画面の回答集計画面から作成する。
4. 報告フォームを作成する。
    * セクション1（ユーザ識別情報および報告日ならびに体温）
      ![report_form_questions_section1.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_section1.png)
    * セクション2（体調不良等の有無）
      ![report_form_questions_section2.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_section2.png)
    * セクション3（体調不良等無し）
      ![report_form_questions_section3.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_section3.png)
    * セクション4（体調不良等有り）
      ![report_form_questions_section4.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_section4.png)
    * セクション1内、体温を尋ねる質問は、プルダウン形式にし、回答に応じてセクションに移動を設定し、37.5℃以上の場合はセクション4（体調不良等有り）に移動させる。
      ![report_form_questions_option1_section1_temparature.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_option1_section1_temparature.png)
    * セクション2内、体調不良等の有無を尋ねる質問は、プルダウン形式にし、回答に応じてセクションに移動を設定し、無しの場合はセクション3（体調不良等無し）に、有りの場合はセクション4（体調不良等有り）に移動させる。
      ![report_form_questions_option2_section2_condition.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_option2_section2_condition.png)
    * セクション3は、それ以降の遷移について、フォームを送信、と設定する。
      ![report_form_questions_option3_section3_submit.png](https://raw.githubusercontent.com/ongakubu/health-report/main/screenshots/report_form_questions_option3_section3_submit.png)
3. 報告フォームの回答を格納するスプレッドシートを、報告フォームの編集画面の回答集計画面から作成する。
4. 登録フォームの編集画面から、スクリプトエディタを起動し、登録フォームに紐づいたApps Scriptプロジェクトを作成する。
5. [register_formフォルダ](https://github.com/ongakubu/health-report/tree/main/register_form)内の各スクリプトを作成する。
6. [register_setup.gs](https://github.com/ongakubu/health-report/blob/main/register_form/register_setup.gs)内部の設定値を定義する。
```
/*健康観察開始日*/
const startDate = new Date();
/*setFullYear(西暦4桁, 月0–11, 日1–31)*/
startDate.setFullYear(2021, 9, 31);  /*例えばこれは2021年10月31日*/
startDate.setHours(4);
startDate.setMinutes(0);
startDate.setSeconds(0);

/*日次メンテナンスの終了時刻*/
const maintenanceTime = new Date();
maintenanceTime.setHours(5);
maintenanceTime.setMinutes(55);

/*締切時刻（演奏会当日など）*/
const deadline  = new Date();
deadline.setFullYear(2021, 10, 20);
deadline.setHours(9);
deadline.setMinutes(0);
deadline.setSeconds(0);
const nMunitutesBeforeForReminder1 = 30;  /*締切時刻30分前のリマインダ*/
const nMunitutesBeforeForReminder2 = 15;  /*締切時刻15分前のリマインダ*/

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
```
7. [sendBulk.gs](https://github.com/ongakubu/health-report/blob/main/register_form/sendBulk.gs)内部にある報告フォームと登録フォームのリンクを先ほど作ったものに置き換える。
```
+ 'https://docs.google.com/forms/d/e/報告フォームのID/viewform?entry.xxxxxxxxx='+studentno+'&entry.xxxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'&entry.xxxxxxxxx='+mailTo+'&entry.xxxxxxxxx='+mobilephone+'&entry.xxxxxxxxx='+todayYear+'-'+todayMonth+'-'+todayDate+'\n\n'
+ 'https://docs.google.com/forms/d/e/登録フォームのID/viewform?entry.xxxxxxxxxx='+studentno+'&entry.xxxxxxxxx='+part+'&entry.xxxxxxxxx='+name+'\n\n'
```
8. 報告フォームの編集画面から、スクリプトエディタを起動し、報告フォームに紐づいたApps Scriptプロジェクトを作成する。
9. [report_formフォルダ](https://github.com/ongakubu/health-report/tree/main/report_form)内の各スクリプトを作成する。
10. [report_setup.gs](https://github.com/ongakubu/health-report/blob/main/report_form/report_setup.gs)内部の設定値を定義する（時刻やファイルIDなどの内容は[register_setup.gs](https://github.com/ongakubu/health-report/blob/main/register_form/register_setup.gs)と一致させる）。
```
/*健康観察開始日*/
const startDate = new Date();
/*setFullYear(西暦4桁, 月0–11, 日1–31)*/
startDate.setFullYear(2021, 9, 31);  /*例えばこれは2021年10月31日*/
startDate.setHours(4);
startDate.setMinutes(0);
startDate.setSeconds(0);

/*日次メンテナンスの終了時刻*/
const maintenanceTime = new Date();
maintenanceTime.setHours(5);
maintenanceTime.setMinutes(55);

/*締切時刻（演奏会当日など）*/
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
```
11. [register_setup.gs](https://github.com/ongakubu/health-report/blob/main/register_form/register_setup.gs)の"initialization"関数を実行する。
12. [report_setup.gs](https://github.com/ongakubu/health-report/blob/main/report_form/report_setup.gs)の"initialization"関数を実行する。
13. 登録フォームのURLを団体構成員に共有して登録してもらう。
14. 健康観察開始日になると配信等が開始される。

## 参考文献・資料
奥村太一（2019）「[Google Appsを用いたオンライン縦断調査システムの構築](https://hdl.handle.net/10513/00007954)」『上越教育大学研究紀要』第38巻第2号、上越教育大学、pp.239-250（参照日：2020年9月14日）。

YEVGENIA（2020）「[GAS(Google Apps Script)ー会社でGAS+Googleフォーム+スプレッドシートを使ってメール通知機能を備えた健康確認アンケートを無料で作成した話](https://blowup-bbs.com/gas-googleform-spreadsheet-helthsheet/)」『Blow Up by Black Swan』2020年4月20日（参照日：2021年3月15日）。

## License
[MIT Licensed.](https://github.com/ongakubu/health-report/blob/main/LICENSE)
