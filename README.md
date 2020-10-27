# COBOLソース解析ツール
"COBOLソース解析ツール"はCOBOLで記述されたソースコードを解析します。<br>
COBOLソースファイルを手動で解析する前に実行することで、コードの目的、処理内容などの全体感を掴むことを目的としています。<br>
解析結果はExcelに記載されます。<br>

# Features
以下の解析結果を出力することができます。
### PGM情報
- COPY句で読み込んでいる定義ファイル名の一覧
- CALL句による呼出モジュール名の一覧
- カーソル一覧
- 使用DB一覧
- 使用DBのSELECT, UPDATEなどCRUDの一覧<br>
※論理名への変換含む<br>
### 関数情報
- 関数の一覧
- 各関数内でのDB操作、および呼出カーソルの一覧
- 各関数から呼び出されている関数の一覧
- 上記呼出関数の内、別モジュールを呼び出している場合はピンク字にて表示
- 到達できない関数はグレーアウトで表示
### 構造図
- 最初に呼び出される関数を起点に、関数の呼出構造を表示<br>
※IF文、EVALUATE文による条件分岐のみ対応
- 構造図内の各ハイパーリンクから対応する関数情報シート内の項目へ遷移可能

# Installation
1. zipファイルをReleasesからダウンロード＋任意の場所で解凍
1. Cドライブ直下に「ソース解析ツール」フォルダを配置<br>
⇒「C:\ソース解析ツール」<br>
※vscodeへの拡張機能追加を行わない場合、ここまでで完了。
1. vscodeから使用する場合、拡張機能メニューを開き、メニュー右上の「…」をクリックし、「VSIXからのインストール」を選択<br><img width="567" alt="1" src="https://user-images.githubusercontent.com/41313415/97319385-2b9dba80-18b0-11eb-9ba8-f5d0137ddb44.PNG">
1. 解凍したフォルダ内にあるvsix拡張子のファイルを選択してインストール<br><img width="470" alt="2" src="https://user-images.githubusercontent.com/41313415/97319904-a7980280-18b0-11eb-94f3-3c1990d9feec.PNG">

# Usage
※以下、A, Bのどちらかを実行する。
## Ａ．vscodeの拡張機能から解析
※特定のソースファイルを解析したい場合に使用
1. 右クリックメニューから「COBOLソース解析」を選んで実行<br>

## Ｂ．windowsフォームアプリから解析
※複数のソースファイルをまとめて解析したい場合に使用<br>
1. 「C:\ソース解析ツール\Cobol_SourceAnalysis\Cobol_SourceAnalysis.exe」を実行<br>
1. フォームが開いたら「ファイル選択」から解析対象となるファイルを指定<br>
1. 「実行」ボタン押下で実行

<br>
上記A, B実行後、<br>
「C:\ソース解析ツール」直下に「COBOLソース解析結果.xlsx」が、<br>
「C:\ソース解析ツール\ExecCobolAnalysis」にログファイルが生成されます。

# Note

このリポジトリから全ての機能をダウンロード可能ですが、"Usage"に記載しているvscodeの拡張機能部分、windowsフォームアプリ部分のソースは別リポジトリにあります。（※このリポジトリではメイン処理のみを管理しています。）<br>
必要に応じてそれぞれ「hane0707/cobol-sourceanalysisextention」、「hane0707/Cobol_SourceAnalysis」を参照してください。

# License
This software is released under the MIT License, see LICENSE.
