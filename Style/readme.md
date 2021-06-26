# 使い方

## FormStyle.cls
### 導入方法
- クラスモジュールへ対象ファイルを追加
- 標準モジュール等 使用したいモジュールのプロシージャ内にいづれかを追記
```VB:UserForm
Private Sub UserForm_Initialize()

  With New FormStyle
    .SetForm Me
    .Style "black"
  End With

End Sub
```
```VB:Module
Private Sub UserForm_Initialize()

  Dim cls As FormStyle
  Set cls = New FormStyle
  With cls
    .SetForm Me
    .Style "black"
  End With
  Set cls = Nothing

End Sub
```
- 処理を実行

### 実装機能

- メソッド

名称|引数|返り値|機能
:-:|:-|:-|:-
SetForm|Form As UserForm = ""|Boolean|Formをセットする
