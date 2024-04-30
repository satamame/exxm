# exxm - EXtract eXcel Macros

## 概要

- 指定のフォルダにある Excel ブックから VBA マクロを抽出します。
- 抽出したマクロを好みのエディタで編集して、元の Excel ブックに戻すことができます。

## 動作環境

- Windows 10 で動作確認しています。
- Excel と .NET Framework 8 がインストールされている必要があります。

## インストール

1. exxm-xxx.zip を解凍して、exxm.exe と exxm-conf.yml を任意のフォルダに配置してください。

## フォルダ構成の例

```
/
├ books/
│ ├ Book1.xlsm
│ └ Book2.xlsm
├ macros/ (最初の抽出時に作られます)
│ ├ Book1.xlsm/
│ │ └ Module1.bas
│ └ Book2.xlsm/
│   └ Module1.bas
├ exxm.exe
└ exxm-conf.yml
```

- books, macros フォルダと exxm-conf.yml は exxm.exe を実行するディレクトリに配置します。
- exxm.exe プログラムは別のディレクトリに置いてパスを指定して実行することもできます。

## 使い方

### 抽出

```
> ./exxm.exe --from-excel
```

### 書き戻し

```
> ./exxm.exe --to-excel
```

## エンコーディング

- 抽出したマクロは Shift-JIS で保存されます。
- VSCode で編集する場合、VBA ファイルタイプを認識する拡張機能をインストールして、settings.json に以下の設定を追加してください。

```json
{
    "[vba]": {
        "files.encoding": "shiftjis"
    }
}
```

## 制限事項

- クラスモジュール、フォームは抽出しません。
- マクロを Excel ブックに書き戻す際は上書きとなります。
    - 上書き対象となるモジュールがない場合はエラーになります。
- Excel のインスタンスが複数起動している場合は抽出や書き戻しができません。
    - これは、どのインスタンスで開いているブックを対象とするか判断できないためです。
    - タスクマネージャで Excel のプロセスを終了してください。
