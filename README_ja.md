# VBAForm2PowerShell - Excel VBA UserForm to PowerShell GUI (WinForms) Converter
🌎[English](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/blob/main/README.md)<br><br>
このプログラムは、Excel VBAにて作成したユーザーフォームをPowerShellのGUI(WinForms)用に変換可能なプログラムです<br>

## 変換例
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/fd6d6445-9dd1-448e-8358-e74f91f571cd" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/46e6fa34-236e-499a-a4fe-a1be7b0e6acc" /><br><br>

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
- コントロールの色(文字色、背景色)
- テキスト表示(Label, CommandButton, CheckBox, ToggleButton, OptionButton, MultiPage)
- フォント(フォント種類、サイズ、太字、斜体)
- 枠線(Frame [Captiionなし], TextBox, Label, ListBox, Image)
- マウスカーソル
- テキスト表示の左寄せ・中央・右寄せ(Label, TextBox, CheckBox, ToggleButton, OptionButton)
- TextBox, ComboBoxのデフォルト値
- ComboBox, ListBoxに設定したアイテム
- OptionButton, CheckBox, ToggleButtonの選択状態
- BackStyleに設定した透明表示設定

## 対応しているコントロールの種類
| VBA Formのクラス | WinFormsのクラス|
| ------ | ------ |
| Label | Label |
| CommandButton | Button |
| Frame (without Caption) | Panel |
| Frame (with any Caption) | GroupBox |
| TextBox | TextBox |
| SpinButton | NumericUpDown |
| ListBox | ListBox |
| CheckBox | CheckBox |
| ToggleButton | CheckBox<br>(Appearance = [System.Windows.Forms.Appearance]::Button) |
| OptionButton | RadioButton |
| Image | PictureBox |
| ScrollBar | HScrollBar / VScrollBar |
| ComboBox | ComboBox |
| MultiPage | TabControl |

※SpinButtonは仕様が異なるため、配置方法によっては外観が異なります<br>
上記以外のコントロールがフォーム上にある場合、変換に失敗するので該当のコントロールを削除したうえで再度変換を行ってください<br>

## 使い方
使用前に、変換したいユーザーフォームが作成されたExcelブックを用意する必要があります<br>
また、VBE上でイミディエイトウィンドウが表示されていない場合は表示の設定を行ってください<br><br>
<img width="843" height="768" alt="Image" src="https://github.com/user-attachments/assets/676cd54c-d610-4c25-bd9a-9e064e38dc5e" /><br><br>
1.[ここ](https://github.com/GUI-Conversion-Tools/VBAForm2PowerShell/releases)から最新版のファイルをダウンロードし解凍してください、中のVBAForm2PowerShell.basを使用します<br>
2. Excelの開発→Visual BasicからVBEを開いてください<br>
3. プロジェクトを右クリックし、「ファイルのインポート」よりVBAForm2PowerShell.basをインポートします<br>
4. イミディエイトウィンドウに「Call ConvertForm2PS(UserForm1)」と入力しEnterキーを押下します<br>
```vb
Call ConvertForm2PS(UserForm1)
```
ダブルクリックで実行可能なbatファイルとして保存したい場合は第二引数をTrueに設定してください<br>
```vb
Call ConvertForm2PS(UserForm1, True)
```
※「UserForm1」の部分は変換したいユーザーフォームのオブジェクト名に変えてください<br>
5. 正常に変換が完了した場合、メッセージが表示されExcelブックと同じディレクトリに「output.ps1」または「output.bat」が作成されます<br>
6. GUIの外観を確認したら、ps1/batファイルを編集し[System.Windows.Forms.Application]::EnableVisualStyles()の上にButton.Add_Click({ 関数名 })でボタン押下時の関数の設定などをしてください<br>

## 子要素を設定できないコントロールの並び順について
WinFormsでは例としてLabelにLabelを重ねた場合は先に設置したコントロールが優先して前面に表示されます<br>
ただしVBAのユーザーフォームにおいては前面/背面を変更することができるためこの限りではありません<br>
このプログラムは各コントロールを逆順に並べ替えた後、階層順にソートして配置します<br>
現状コントロールのZオーダー(前面/背面情報)を取得できる手段がないため反映させることができずVBAでの表示と異なってしまう場合があります<br>
その場合は、PowerShellのコードを編集し、.BringToFront() または .SendToBack()を使用し調整を行ってください<br>
なお、新規でGUIを作成する場合は重ねるよりもFrameなどの明確な親子関係を設定可能なコントロールを使用することを推奨します<br>

## 使用のさいの注意点
マルチモニター環境でこのプログラムを使用する場合、一時的にモニターを1つにするか、すべてのモニターの拡大率を統一したうえで使用してください<br>
異なる拡大率のモニターが混在している場合、ウィンドウサイズの計算が正常に行えない可能性があります<br>
