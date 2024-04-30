Attribute VB_Name = "Module1"
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SlideByTransWriter
'第１言語の後に第２言語をワンクリック（orショートカット）で併記するアドイン．
'一番下までスクロールして，「====入力欄====」パラメータを編集することで，フォントサイズなどを指定できる．
'リボンにフォームを作ったが，一番左のボタン以外機能しないので注意．フォームからテキストを取得するコールバックがかけなかったので，有志求ム．
'DeeplのAPIキーが間違っているとフリーズします．
'
'ライブラリリファレンス：以下URLからPowerPoint参照
'[Visual Basic for Applications (VBA) のライブラリ リファレンス | Microsoft Learn](https://learn.microsoft.com/ja-jp/office/vba/api/overview/library-reference)
'
'対応言語一覧
'[OpenAPI spec for text translation | English | DeepL API Docs](https://developers.deepl.com/docs/api-reference/translate/openapi-spec-for-text-translation)
'全体を通して，翻訳前をsource,翻訳後をtargetとする
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'Sleep文を使うためのおまじない


'現在選択しているテキストボックスのテキストを取得
Function get_select_shape_text() As String

    With ActiveWindow.Selection.ShapeRange '現在選択している図形（テキストボックス含む）
    
        Dim source_text As String: source_text = .TextFrame2.TextRange.Text 'テキストを取得
        Debug.Print source_text
        get_select_shape_text = source_text
        
    End With
End Function

'現在選択しているtextboxの下にある大きさのtarget-textboxを生成する
Function create_target_textbox(target_text As String, indent_width As Double, target_textbox_height As Double, tb_bold As Boolean, tb_color As Long, tb_size As Double, tb_fontname As String)

    With ActiveWindow.Selection.ShapeRange '現在選択している図形（textbox含む）
    
        'target-textboxのLeftはsource-textboxのLeft+インデント幅
        Dim target_textbox_left As Double: target_textbox_left = .Left + indent_width
        'target-textboxのTopはsource-textboxの下端
        Dim target_textbox_top As Double: target_textbox_top = .Top + .Height
        'target-textboxのWidthはsource-textboxと同じ
        Dim target_textbox_width As Double: target_textbox_width = .width
        
        Set active_window = ActiveWindow.Selection.SlideRange '生成の都合上，現在アクティブなスライドのオブジェクトを得ておく
        
        'target-textboxを生成
        Dim target_textbox As Shape: Set target_textbox = active_window.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=target_textbox_left, Top:=target_textbox_top, width:=target_textbox_width, Height:=target_textbox_height) 'AddTextbox(文字の向き，左上点x座標，左上点y座標,ボックスの幅，高さ)
        
        'Debug.Print tb_bold
        With target_textbox.TextFrame.TextRange
            .Text = target_text '.TextFrame.TextRange.Text ="内容"
            .Font.Bold = tb_bold
            .Font.Color.RGB = tb_color
            .Font.Size = tb_size
            .Font.Name = tb_fontname
        End With
        
    End With

End Function

Function create_shell_arg(arg As String) As String
    Dim shell_arg As String: shell_arg = Chr$(34) & arg & Chr$(34) & " "
    create_shell_arg = shell_arg
End Function

'翻訳
Function translate(source_text As String, source_lang As String, target_lang As String, api_key As String, exe_folder As String) As String

    ' Shell関数を使用してPythonスクリプトを実行
    Dim translate_cmd As String: translate_cmd = create_shell_arg(exe_folder & "\translator.exe") & create_shell_arg(source_text) & create_shell_arg(source_lang) & create_shell_arg(target_lang) & create_shell_arg(api_key) & create_shell_arg(exe_folder) ' xxx.exe 引数1 引数2...の形式のコマンド
    Debug.Print translate_cmd 'コマンド内容確認
    Shell translate_cmd, vbHide 'コマンド発動
    
    '翻訳語文書のテキストファイルが格納される予定のパス
    Dim target_text_path As String
    target_text_path = exe_folder & "\translated.txt"
    'Debug.Print target_text_path
    
    '翻訳語文書のテキストファイルが生成されてないならば，まだ翻訳終わってないので待つ
    Do While Dir(target_text_path) = ""
        'Debug.Print "wait"
        Sleep 10 'このスリープがないと，虚無while許さないマンによりエラー起こる
    Loop
        
    Dim target_text As String '翻訳語文書
    
    ' 翻訳語文書のテキストファイルから読み込み
    Open target_text_path For Input As #1
            Line Input #1, target_text
    Close #1
    
    Debug.Print target_text '翻訳語文書の確認
    'Debug.Print target_text_path
    Kill target_text_path ' 翻訳語文書のテキストファイルを消去　このファイルの有無で翻訳待ちを判断するので，消しておく
    
    translate = target_text

End Function


'メイン関数みたいなもん
Sub SlideByTransWriter()

    '===============================入力欄===========================================================================================================================================================
    
    '翻訳言語設定値 / Translation language setpoint / Example Japanese:JA, English:EN
    Dim source_lang As String: source_lang = "JA" '翻訳前言語 / pre-translated language
    Dim target_lang As String: target_lang = "EN" '翻訳後言語 / post-translational language
    
    '翻訳サービスまわりの設定値 / Set values around translation services
    Dim api_key As String: api_key = "xxxxxxxxxxxxxx" 'DeepL API Key
    Dim exe_folder As String: exe_folder = "yyyyyyyyyyyyyyyyy" 'translator.exeのあるフォルダ / Folder with translator.exe
    
    'TextBoxのフォントの設定値 / Font setting value for TextBox
    Dim tb_bold As Boolean: tb_bold = False '太字にするならTrue,そうでないならFalse
    Dim tb_color As Long: tb_color = RGB(166, 166, 166) 'フォントの色
    Dim tb_size As Double: tb_size = 20 'フォントサイズ
    Dim tb_fontname As String: tb_fontname = "Arial"
    
    'TextBoxの位置や大きさの設定値 / Set values for TextBox position and size
    Dim indent_width As Double: indent_width = 50 'インデント幅[pt] / indent width[pt]
    Dim target_textbox_height As Double: target_textbox_height = 40 '翻訳後テキストボックスの高さ[pt]．フォントサイズより大きくすること．/ Height of the text box after translation [pt]. Should be larger than the font size.
    
    '===============================入力欄===========================================================================================================================================================
        
    Dim source_text As String: source_text = get_select_shape_text() '現在選択しているtextboxのテキスト=source-textを取得
    
    Dim target_text As String: target_text = translate(source_text, source_lang, target_lang, api_key, exe_folder) '翻訳し，target-textを取得
    
    Call create_target_textbox(target_text, indent_width, target_textbox_height, tb_bold, tb_color, tb_size, tb_fontname) '現在選択しているtextboxの下にある大きさのtarget-textboxを生成する
    
End Sub






