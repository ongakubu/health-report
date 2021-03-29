# health-report
新型コロナウイルス感染症（COVID-19）対応の一環として、[学習院輔仁会音楽部](https://www.ongakubu.org)において開発・運用している健康観察システムです。
本システムを用いて、対面活動参加部員が日々行う検温や体調の報告・集計を自動化しています。
Google Forms、Spreadsheet、Apps Scriptを用いており、質問項目等は学生オーケストラでの利用を想定したものになっていますから、必要に応じて変更・調整してご利用ください。

## 主な機能
* 報告フォーム（体温、体調等）をメールで定時配信
* ユーザ識別情報のフォームへの事前入力
* 37.5℃以上の発熱ないし体調不良発生時は管理者に自動で通知

## 仕組み
![20210329_health_report_system_v3](https://user-images.githubusercontent.com/18068336/112787095-bae31000-9092-11eb-90c1-cdf4764f5e87.jpg)

1. ユーザが登録フォームに必要事項を入力して送信する。
2. 登録フォームが登録確認メールをユーザに自動送信する。
3. 登録されたユーザのメールアドレスが宛先一覧スプレッドシートに追加される。
4. 定時になると報告フォームへのリンクを含むメールが生成される。
5. 生成されたメールが宛先一覧に記載されているメールアドレスに自動配信される。
6. ユーザが配信された報告フォームに必要事項（体温、体調等）を入力して回答を送信する。
7. もし発熱や体調不良等が報告された場合は、管理者にメールで自動通知する。

## 構築方法
### 1. 登録フォームを作成する。
#### 1.1. 質問を作成する。
* [登録フォーム質問項目](https://github.com/ongakubu/health-report/blob/main/register_form/register_form_questions.md)
#### 1.2. 登録フォームの回答を保存するスプレッドシートを作成する。
登録フォームの回答編集画面からスプレッドシートを作成する。
* [登録フォーム回答](https://github.com/ongakubu/health-report/blob/main/register_form_answer/register_form_answer.md)
#### 1.3. 登録フォームに自動返信スクリプトを追加する。
登録フォームの編集画面からスクリプトエディタを開き、次のプログラムを追加し、変数を適切に設定する。
* [confirmRegistration.gs](https://github.com/ongakubu/health-report/blob/main/register_form/confirmRegisteration.gs)
#### 1.4. 登録フォームの自動返信スクリプトにトリガーを設定する。
スクリプトエディタのトリガー編集画面にて、トリガーを設定する。
|No.|実行する関数を選択|デプロイ時に実行|イベントのソースを選択|イベントの種類を選択|エラー通知設定|
|:--|:--|:--|:--|:--|:--|
|1|sendRegisterConfirm|Head|フォームから|フォーム送信時|今すぐ通知を受け取る|
### 2. 宛先一覧スプレッドシートを作成する。
#### 2.1. 空のスプレッドシートを作成する。
* [recipients](https://github.com/ongakubu/health-report/blob/main/recipients/recipients.md)
#### 2.2 宛先一覧スプレッドシートにスクリプトを追加する。
メニューバー＞ツール＞スクリプトエディタを開き、次のプログラムを追加し、変数を適切に設定する。
* [manageRecipientsSendBulk.gs](https://github.com/ongakubu/health-report/blob/main/recipients/manageRecipientsSendBulk.gs)
#### 2.3. 宛先一覧スプレッドシートのスクリプトにトリガーを設定する。
スクリプトエディタのトリガー編集画面にて、次の3つのトリガーを設定する。
|No.|実行する関数を選択|デプロイ時に実行|イベントのソースを選択|時間ベースのトリガーのタイプを選択|時刻を選択|エラー通知設定|
|:--|:--|:--|:--|:--|:--|:--|
|1|getRecipients|Head|時間主導型|特定の日時|健康観察開始日時を設定|今すぐ通知を受け取る|
|2|addRecipients|Head|時間主導型|日付ベースのタイマー|メール配信時刻の1時間前|今すぐ通知を受け取る|
|3|sendEveryoneEveryday|Head|時間主導型|日付ベースのタイマー|メール配信時刻（例えば午前6時〜7時）|今すぐ通知を受け取る|
### 3. 報告フォームを作成する。
#### 3.1. 質問を作成する。
* [報告フォーム質問項目](https://github.com/ongakubu/health-report/blob/main/report_form/report_from_questions.md)
#### 3.2. 報告フォームの回答を保存するスプレッドシートを作成する。
報告フォームの回答編集画面からスプレッドシートを作成する。
* [報告フォーム回答](https://github.com/ongakubu/health-report/blob/main/report_form_answer/report_form_answer.md)
#### 3.3. 報告フォームに自動返信スクリプトを追加する。
報告フォームの編集画面からスクリプトエディタを開き、次のプログラムを追加し、変数を適切に設定する。
* [notifyManagers.gs](https://github.com/ongakubu/health-report/blob/main/report_form/notifyManagers.gs)
#### 3.4. 報告フォームの自動返信スクリプトにトリガーを設定する。
スクリプトエディタのトリガー編集画面にて、トリガーを設定する。
|No.|実行する関数を選択|デプロイ時に実行|イベントのソースを選択|イベントの種類を選択|エラー通知設定|
|:--|:--|:--|:--|:--|:--|
|1|notifyManagers|Head|フォームから|フォーム送信時|今すぐ通知を受け取る|

## 参考文献・資料
奥村太一（2019）「[Google Appsを用いたオンライン縦断調査システムの構築](https://hdl.handle.net/10513/00007954)」『上越教育大学研究紀要』第38巻第2号、上越教育大学、pp.239-250（参照日：2020年9月14日）。

YEVGENIA（2020）「[GAS(Google Apps Script)ー会社でGAS+Googleフォーム+スプレッドシートを使ってメール通知機能を備えた健康確認アンケートを無料で作成した話](https://blowup-bbs.com/gas-googleform-spreadsheet-helthsheet/)」『Blow Up by Black Swan』2020年4月20日（参照日：2021年3月15日）。

## 免責事項
この健康観察システムを利用した結果生じたあらゆる損害等に対しても、学習院輔仁会音楽部は理由を問わず一切責任を負いません。
