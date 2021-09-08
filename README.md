# health-report
新型コロナウイルス感染症（COVID-19）対応の一環として、[学習院輔仁会音楽部](https://www.ongakubu.org)において開発・運用している健康観察システムです。
本システムを用いて、対面活動参加部員が日々行う検温や体調の報告・集計を自動化しています。
Google Forms、Spreadsheet、Apps Scriptを用いており、フォームの質問項目等は学生オーケストラでの利用を想定したものになっていますから、必要に応じて変更・調整してご利用ください。

## 主な機能
* 新規登録・登録内容更新・登録解除・報告済健康観察記録の照会手続をフォームから自動処理
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

## 参考文献・資料
奥村太一（2019）「[Google Appsを用いたオンライン縦断調査システムの構築](https://hdl.handle.net/10513/00007954)」『上越教育大学研究紀要』第38巻第2号、上越教育大学、pp.239-250（参照日：2020年9月14日）。

YEVGENIA（2020）「[GAS(Google Apps Script)ー会社でGAS+Googleフォーム+スプレッドシートを使ってメール通知機能を備えた健康確認アンケートを無料で作成した話](https://blowup-bbs.com/gas-googleform-spreadsheet-helthsheet/)」『Blow Up by Black Swan』2020年4月20日（参照日：2021年3月15日）。

## License
[MIT Licensed.](https://github.com/ongakubu/health-report/blob/main/LICENSE)
