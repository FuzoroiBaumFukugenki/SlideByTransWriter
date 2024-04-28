# SlideByTransWriter
PowerPointでテキストボックスの真下に，翻訳文入りのテキストボックスをワンクリックで生成．多言語併記を助けます．

https://github.com/FuzoroiBaumFukugenki/SlideByTransWriter/assets/147458822/9f80f6a4-daa1-4f4e-bf47-9db374cf7b49

<div align="center">

  <p style="display: inline">
    <img src="https://custom-icon-badges.herokuapp.com/badge/license-GPL%203.0-8BB80A.svg?logo=law&logoColor=white">
    <img src="https://custom-icon-badges.herokuapp.com/badge/VBA-867db1.svg?logo=VBA&logoColor=white">
    <img src="https://custom-icon-badges.herokuapp.com/badge/Shell-89e051.svg?logo=Shell&logoColor=white">
    <img src="https://img.shields.io/badge/windows-0078D6?logo=windows&logoColor=white">
    <img src="https://img.shields.io/badge/Python-14354C.svg?logo=python&logoColor=white">
    <img src="https://img.shields.io/badge/Microsoft PowerPoint-B7472A?logo=microsoft-powerpoint&logoColor=white">

    
  </p>

  [日本語版](#日本語版)

</div>

# 日本語版


## 要件
- windows10もしくは11
- Microsoft 365およびPowerPointが導入済み


## セットアップ

### 翻訳するソフト（.exe）の設置

1. 「SlideByTransWriter」フォルダごと任意の場所に設置．（※クラウド非推奨）
2. このとき，設置したフォルダのパスをメモ帳などに記録しておく．
   - 以降，このパスをTRANSLATOR_PATHとする．
   - 例：TRANSLATOR_PATH = "C:/slideByTransWriter_Translator"(translator.exeは含まない)
   - 「TRANSLATOR_PATH/translator.exe」が日本語を入力に英語を出力してくれるソフトです．Pythonでできてます．詳細は

### マクロを動作できるようにする＆翻訳文のフォントを変える

1. SlideByTransWriter.pptmを右クリックし，プロパティを開く
2. 「全般>セキュリティ」から「許可する」にチェックを入れ，「適用」をクリック．
    - マクロを実行するために必要です．これが許諾できない場合，あきらめてください．
    - 人によっては必要ない項目です
3. SlideByTransWriter.pptmを開く．
    - 編集を許可してください
5. 「開発タブ>Visual Basic」よりVBAのエディタを開く．
    - 開発タブがない場合，「ファイルタブ>オプション>リボンのユーザー設定」より「開発」にチェックを入れてください
7. VBAのエディタの左側のプロジェクトペインから「標準モジュール」を展開，「Module1」を開きSlideByTransWriterのVBAマクロが記述されていることを確認．
8. ここでVBAマクロの一番下までスクロールし，変数の中身を変える形でDeepLのAPIキーとTRANSLATOR_PATHを入力する．また，同様にして翻訳文のフォントの変えることができる．
    - DeepLのAPIキーの入手はこちらなどを参考にしてください．[DeepL API用のAPIキー – DeepLヘルプセンター | どんなことでお困りですか？](https://support.deepl.com/hc/ja/articles/360020695820-DeepL-API%E7%94%A8%E3%81%AEAPI%E3%82%AD%E3%83%BC)
9. 「こんにちは」など文書を選択した状態で，再生ボタンマークからVBAを実行．SlideByTransWriterの動作確認をする．
10. 満足のいくフォントに調整し終わるまで，6,7を繰り返す．
11.  以降「開発タブ>マクロ>SlideByTransWriter>実行」より都度翻訳できるようになる．

### SlideByTransWriterのアドイン化

- 毎度「開発タブ>マクロ>SlideByTransWriter>実行」して翻訳するのは手間な上このファイル上っでしか翻訳できない．そこで，SlideByTransWriterをアドイン化して
  - ワンクリックもしくはショートカットキーで翻訳できるようにする．
  - 新規にPowerPointファイルを作った時に，すぐにSlideByTransWriterを使えるようにする．

1. 「ファイルタブ>名前を付けて保存>参照」から「.ppam」（パワーポイントアドイン）の形式を指定．
    - エクスプローラーが開いた後の「ファイルの種類」から.ppamが選択可能です．
2. 自動で「C:\Users\ [User]\AppData\Roaming\Microsoft\AddIns」が表示されるので，そこに「SlideByTransWriter.ppam」の名前で保存
3. 新規にPowerPointファイルを作る．「開発タブ>PowerPointアドイン>新規追加」から先ほど追加した「SlideByTransWriter.ppam」を選択
4. リボンのタブに「SlideByTransWriter」が追加されていることを確認．
5. ”こんにちは”など任意の日本語が書かれたテキストボックスを選択後「SlideByTransWriterタブ>Insert Translated text」をクリックすると翻訳されることを確認．
    - ここからはデモ映像の最後の方に収録されています．
    - 言語選択タブなどありますが，Insert Translated textボタン以外動作しませんので注意してください．
6. 「SlideByTransWriterタブ>Insert Translated text」を右クリックし「クイックアクセスツールバーに追加」を選択．
7. Altキーを押すとショートカットのガイドが出てくるので，それに応じてショートカットで翻訳できることを確認する．
    - 開発者の環境（デモ映像）だとAlt+8キーでしたが，個人ごと既存のクイックアクセスツールなどにより変わってきます．

この設定をしておくことで，次に新規作成するPowerPointファイルにはじめから，SlideByTransWriterタブとクイックアクセスツールが常駐します．

<p align="right"><a href="#top">Back to TOP</a></p>

## 手の加え方

### 開発環境
少なくともこれで動いて開発しました．
モジュールは特別インストールが必要だったもののみ記載してます．

| 言語・フレームワークなど  | バージョン |
| --------------------- | ---------- |
| Python                | 3.11.9     |
| (Python)requests   |  2.31.0     |
| (Python)Pyinstaller   | 6.6.0    |
| DeepL API | 2.11.0     |
| Windows               | Windows 11 Home 23H2 |

### 翻訳部分（Python）
[開発環境](#開発環境)をそろえて「translator.py」を編集後，pyinstallerで「translator.exe」という名前でTRANSLATOR_PATHに設置してください．(pyinstaller以外のexe化は担保できないです．)翻訳コードの詳細はtranslator.py内のコメントに任せます．

### PowerPointの挙動（VBA,CustomUI）

#### VBA
[マクロを動作できるようにする＆翻訳文のフォントを変える](#マクロを動作できるようにする＆翻訳文のフォントを変える)の5番工程から編集可能です．コードの詳細はコメントに任せます．

#### CustomUI
Office RibbonX EditorからXMLを編集する形でリボンにあるボタンを作ってます．以下のサイト様の解説が分かりやすいです．

[VBAからPowerPointアドインを自作する方法](https://zenn.dev/mtsuda/articles/7404280eb5dbb2#fn-7efb-3)

### TODO

- VBAで直接APIが叩けるようにする．そうするとPythonが要らなくなる．
- リボンに言語選択タブなど付けたが，VBA内でその値を取得する方法（=コールバック関数）の書き方が分からない．理想はAPIキーなど設定値をすべてリボンで設定したい．

<p align="right"><a href="#top">Back to TOP</a></p>

## トラブルシューティング

### 後から翻訳後テキストボックスのフォントを変えたい．
[マクロを動作できるようにする＆翻訳文のフォントを変える](#マクロを動作できるようにする＆翻訳文のフォントを変える)の８番の工程からやり直してください．再アドイン化しないと適用されないので，[SlideByTransWriterのアドイン化](#SlideByTransWriterのアドイン化)もその後きちんと行ってください．

### Selection（無効なメンバー）無効な要求です．適切な項目が選択されていません．

テキストボックスなどのシェイプを選択できていません．選択状態でマクロを開始してください．

### フリーズする

（エラー処理書いていないので）APIキーが間違ってるか利用不可だとフリーズします．APIキーの再確認をしてください．

### これ以上読み込める内容がありません．

日本語と英語以外は確認できてないので，文字コードエラーが起こってる可能性が高いです．希望があれば対応します．

### 何も起こらない．もしくは「マクロが実行できません」系統のエラーが出る

「ファイル>オプション>トラストセンター>トラストセンターの設定」より「すべてのマクロを有効にする」を選択してください．

### 英語から英語，日本語から日本語が出てくる．

翻訳言語の指定が誤っています．[マクロを動作できるようにする＆翻訳文のフォントを変える](#マクロを動作できるようにする＆翻訳文のフォントを変える)の８番の工程から翻訳言語を変更してください．

<p align="right"><a href="#top">Back to TOP</a></p>
