# health-report
新型コロナウイルス感染症（COVID-19）対応の一環として、[学習院輔仁会音楽部](https://www.ongakubu.org)において開発・運用している健康観察システムです。
本システムを用いて、対面活動参加部員が日々行う検温や体調の報告・集計を自動化しています。
Google Forms、Spreadsheet、Apps Scriptを用いており、フォームの質問項目等は学生オーケストラでの利用を想定したものになっていますから、必要に応じて変更・調整してご利用ください。

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

## 参考文献・資料
奥村太一（2019）「[Google Appsを用いたオンライン縦断調査システムの構築](https://hdl.handle.net/10513/00007954)」『上越教育大学研究紀要』第38巻第2号、上越教育大学、pp.239-250（参照日：2020年9月14日）。

YEVGENIA（2020）「[GAS(Google Apps Script)ー会社でGAS+Googleフォーム+スプレッドシートを使ってメール通知機能を備えた健康確認アンケートを無料で作成した話](https://blowup-bbs.com/gas-googleform-spreadsheet-helthsheet/)」『Blow Up by Black Swan』2020年4月20日（参照日：2021年3月15日）。

## License
[MIT Licensed.](https://github.com/ongakubu/health-report/blob/main/LICENSE)
