# myBridge Organizer
名刺管理ツール[myBridge](jp.mybridge.com)からエクスポートした複数のExcel名刺帳を統合して出力するプログラムです。

## Input
引数で指定したExcelを読み込みデータベースに記録します。
過去に読み込んだデータは基本的に削除されず、蓄積され続けます。

Excelの各行を読み込むかどうかの判断フローは以下の通りです。
1. 読み込み対象の行と「会社名」「名前」「部署」のすべてが一致するデータがデータベースにあるかを調べます。
1. もしすべてが一致するデータがなければ読み込みます。
1. すべてが一致するデータがあった場合、「名刺交換日」が新しい方のデータのみを採用します。

## Output
データベースのデータをExcelとGitHub Issueに出力します。

Excelへの出力は`output`フォルダに会社名ごとのファイルに分けて出力します。

GitHub Issueへの出力は https://github.com/shiroi36/Drawing/issues/547 にコメントとして出力します。
出力する前に過去に出力したコメントはその都度すべて削除します。
