# VBAForm2PowerShell - Excel VBA UserForm to PowerShell GUI (WinForms) Converter
🌎[English](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/blob/main/README.md)<br><br>
このプログラムは、Excel VBAにて作成したユーザーフォームをPowerShellのGUI(WinForms)用に変換可能なプログラムです<br>

## 変換例
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/6762d0bb-16f5-4555-8700-95e3c4db19d5" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/98cb0e04-e541-447d-9ff6-18469dcfe49f" /><br><br>

## 動作要件
- 対応OS: Windows
- 必要ソフトウェア: Microsoft Excel

## 動作確認済環境
- Windows 10/11
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)

## 反映する項目
- 変数名(オブジェクト名)
- コントロールのおおよそのレイアウトとサイズ
- コントロールの色(文字色) (ただし次のコントロールを除く: `MultiPage`, `ComboBox` [.Style = fmStyleDropDownList])
- コントロールの色(背景色) (ただし次のコントロールを除く: `MultiPage`, `ComboBox` [.Style = fmStyleDropDownList], `ScrollBar`)
- テキスト表示(`Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- フォント(フォント種類、サイズ、太字、斜体)
- 枠線(`Frame [Captionなし]`, `TextBox`, `Label`, `ListBox`, `Image`)
- マウスカーソル
- テキスト表示の左寄せ・中央・右寄せ(`Label`, `TextBox`, `CheckBox`, `ToggleButton`, `OptionButton`)
- `TextBox`, `ComboBox`のデフォルト値
- `ComboBox`, `ListBox`に設定したアイテム
- `OptionButton`, `CheckBox`, `ToggleButton`の選択状態
- `.BackStyle`に設定した透明表示設定 (ただし次のコントロールを除く: `ComboBox` [.Style = fmStyleDropDownList])
- `.Orientation`/`.Min`/`.Max`プロパティ (`ScrollBar`)
- `.Alignment`プロパティ (`CheckBox`. `OptionButton`)
- `.TabOrientation`プロパティ (`MultiPage`)
- `.Locked`プロパティ (`TextBox`, `ListBox`, `ComboBox`)
- `.PasswordChar`プロパティ (`TextBox`)
- `.Style`プロパティ (`ComboBox`, `MultiPage`)
- `.MultiSelect`プロパティ (`ListBox`)
- `.PictureAlignment`/`.PictureSizeMode`プロパティ (`Image`)

>注:
>
>-   `.BackStyle`が `fmBackStyleOpaque`の場合VBAのコントロールの`.BackColor`を直接適用します
>-   `.BackStyle`が `fmBackStyleTransparent`の場合:
>        -   WinFormsが透明色をサポートしているコントロール(例: `Label`, `CommandButton`, `CheckBox`, `OptionButton`, `Image`など)の場合`.BackColor = "Transparent"`に設定します
>        -   WinFormsが透明色をサポートしていないコントロール(`TextBox`, `ComboBox`, `ToggleButton`)では以下のように変換されます:
>            -   親コントロールが`.BackColor`プロパティを持つ場合、その色を`.BackColor`に設定します
>            -   親コントロールが`Page`の場合、`.BackColor`プロパティを持たないため`Page`の視覚的な背景色と一致するシステムカラーの`&H8000000F&`を`.BackColor`に設定します
>
>-   `.PictureSizeMode`/`.PictureAlignment`は対応するWinFormsの `.SizeMode`に変換されます:
>        -   `fmPictureSizeModeClip` → `"Normal"` または `"CenterImage"` (`.PictureAlignment`の設定値により変化)
>        -   `fmPictureSizeModeStretch` → `"StretchImage"`
>        -   `fmPictureSizeModeZoom` → `"Zoom"`
>
>        -   `.PictureSizeMode = fmPictureSizeModeClip`の場合:
>            -   `.PictureAlignment = fmPictureAlignmentCenter` → `"CenterImage"`
>            -   `.PictureAlignment = fmPictureAlignmentTopLeft` → `"Normal"`
>            -   WinFormsの`PictureBox`は左上と中央配置のみをサポートしているため、他の値は`"Normal"`(左上)に変換されます
>        -   `.PictureSizeMode`が`fmPictureSizeModeStretch`または`fmPictureSizeModeZoom`の場合`.PictureAlignment`は無視されます
>-   MultiPageコントロールの`.TabOrientation`が`fmTabOrientationLeft`または`fmTabOrientationRight`になっている場合タブのテキストが縦方向になり、横方向を維持するVBAとは見た目が異なります

## 対応しているコントロールの種類
| VBA Formのクラス | WinFormsのクラス|
| ------ | ------ |
| `Label` | `Label` |
| `CommandButton` | `Button` |
| `Frame` (Captionなし) | `Panel` |
| `Frame` (Captionあり) | `GroupBox` |
| `TextBox` | `TextBox` |
| `SpinButton` | `NumericUpDown` |
| `ListBox` | `ListBox` |
| `CheckBox` | `CheckBox` |
| `ToggleButton` | `CheckBox`<br>(`.Appearance = "Button"`) |
| `OptionButton` | `RadioButton` |
| `Image` | `PictureBox` |
| `ScrollBar` | `HScrollBar` / `VScrollBar` |
| `ComboBox` | `ComboBox` |
| `MultiPage` | `TabControl` |

※`SpinButton`は仕様が異なるため、配置方法によっては外観が異なります<br>
上記以外のコントロールがフォーム上にある場合、変換に失敗するので該当のコントロールを削除したうえで再度変換を行ってください<br>

## 使い方
使用前に、変換したいユーザーフォームが作成されたExcelブックを用意する必要があります<br>
また、VBE上でイミディエイトウィンドウが表示されていない場合は表示の設定を行ってください<br><br>
<img width="807" height="768" alt="Image" src="https://github.com/user-attachments/assets/b023597f-6f9e-4223-a9a4-1c7c499c194b" /><br><br>
1.[ここ](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/releases)から最新版のファイルをダウンロードし解凍してください、中のVBAForm2PowerShell.basを使用します<br>
2. Excelの開発→Visual BasicからVBEを開いてください<br>
3. プロジェクトを右クリックし、「ファイルのインポート」より`VBAForm2PowerShell.bas`をインポートします<br>
4. イミディエイトウィンドウに`Call ConvertForm2PS(UserForm1)`と入力しEnterキーを押下します<br>
```vb
Call ConvertForm2PS(UserForm1)
```
ダブルクリックで実行可能なbatファイルとして保存したい場合は第二引数を`True`に設定してください<br>
```vb
Call ConvertForm2PS(UserForm1, True)
```
※`UserForm1`の部分は変換したいユーザーフォームのオブジェクト名に変えてください<br>
5. 正常に変換が完了した場合、メッセージが表示されExcelブックと同じディレクトリに`output.ps1`または`output.bat`が作成されます<br>
6. GUIの外観を確認したら、`.ps1`/`.bat`ファイルを編集し`.ShowDialog()`の上に`Button.Add_Click({ 関数名 })`でボタン押下時の関数の設定などをしてください<br>

## 引数

`ConvertForm2PS` には以下の引数を設定できます:

|**引数**|**型**|**説明**|
|----------------|-------------------------------|-----------------------------|
|`frms` |`Variant`|**必須**<br>変換対象の`UserForm`オブジェクトまたは`UserForm`オブジェクトの配列を指定 |
|`saveAsBat` |`Boolean`|**省略可能 (デフォルト: `False`).**<br>`True`にした場合PowerShellスクリプトをダブルクリックで起動可能な`.bat`ファイルとして保存する|
|`useCls`  |`Boolean` |**省略可能 (デフォルト: `False`).**<br>`True`にした場合生成したPowerShellコードにおいて各フォームをクラス化する&nbsp;&nbsp;`frms`が配列の場合は自動的に`True`に設定される|
|`noMainLoop`  |`Boolean`|**省略可能 (デフォルト: `False`).**<br>`True`にした場合生成したPowerShellスクリプトに`.ShowDialog()`を含めなくする &nbsp;&nbsp;`useCls`が`True`の場合はインスタンスの作成(例:`$obj_UserForm1 = [UserForm1]::new()`)もスキップする|

`ConvertForm2PS`は単一のユーザーフォームまたは配列内の複数のユーザーフォームを変換することが可能です

```vb
' 実行例: 単一のフォームを変換
Call ConvertForm2PS(UserForm1)

' 実行例: 単一のフォームを変換 (クラス化を行う)
Call ConvertForm2PS(UserForm1, useCls:=True)

' 実行例: 複数のフォームを変換 (自動的にクラス化される)
Call ConvertForm2PS(Array(UserForm1, UserForm2))
```

## 子要素を設定できないコントロールの並び順について
WinFormsでは例として`Label`に`Label`を重ねた場合は先に設置したコントロールが優先して前面に表示されます<br>
ただしVBAのユーザーフォームにおいては前面/背面を変更することができるためこの限りではありません<br>
このプログラムは各コントロールを逆順に並べ替えた後、階層順にソートして配置します<br>
現状コントロールのZオーダー(前面/背面情報)を取得できる手段がないため反映させることができずVBAでの表示と異なってしまう場合があります<br>
その場合は、PowerShellのコードを編集し、`.BringToFront()`または`.SendToBack()`を使用し調整を行ってください<br>
なお、新規でGUIを作成する場合は重ねるよりも`Frame`などの明確な親子関係を設定可能なコントロールを使用することを推奨します<br>

## 使用のさいの注意点
マルチモニター環境でこのプログラムを使用する場合、一時的にモニターを1つにするか、すべてのモニターの拡大率を統一したうえで使用してください<br>
異なる拡大率のモニターが混在している場合、ウィンドウサイズの計算が正常に行えない可能性があります<br>
