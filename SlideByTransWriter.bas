Attribute VB_Name = "Module1"
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SlideByTransWriter
'��P����̌�ɑ�Q����������N���b�N�ior�V���[�g�J�b�g�j�ŕ��L����A�h�C���D
'��ԉ��܂ŃX�N���[�����āC�u====���͗�====�v�p�����[�^��ҏW���邱�ƂŁC�t�H���g�T�C�Y�Ȃǂ��w��ł���D
'���{���Ƀt�H�[������������C��ԍ��̃{�^���ȊO�@�\���Ȃ��̂Œ��ӁD�t�H�[������e�L�X�g���擾����R�[���o�b�N�������Ȃ������̂ŁC�L�u�����D
'Deepl��API�L�[���Ԉ���Ă���ƃt���[�Y���܂��D
'
'���C�u�������t�@�����X�F�ȉ�URL����PowerPoint�Q��
'[Visual Basic for Applications (VBA) �̃��C�u���� ���t�@�����X | Microsoft Learn](https://learn.microsoft.com/ja-jp/office/vba/api/overview/library-reference)
'
'�Ή�����ꗗ
'[OpenAPI spec for text translation | English | DeepL API Docs](https://developers.deepl.com/docs/api-reference/translate/openapi-spec-for-text-translation)
'�S�̂�ʂ��āC�|��O��source,�|����target�Ƃ���
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'Sleep�����g�����߂̂��܂��Ȃ�


'���ݑI�����Ă���e�L�X�g�{�b�N�X�̃e�L�X�g���擾
Function get_select_shape_text() As String

    With ActiveWindow.Selection.ShapeRange '���ݑI�����Ă���}�`�i�e�L�X�g�{�b�N�X�܂ށj
    
        Dim source_text As String: source_text = .TextFrame2.TextRange.Text '�e�L�X�g���擾
        Debug.Print source_text
        get_select_shape_text = source_text
        
    End With
End Function

'���ݑI�����Ă���textbox�̉��ɂ���傫����target-textbox�𐶐�����
Function create_target_textbox(target_text As String, indent_width As Double, target_textbox_height As Double, tb_bold As Boolean, tb_color As Long, tb_size As Double, tb_fontname As String)

    With ActiveWindow.Selection.ShapeRange '���ݑI�����Ă���}�`�itextbox�܂ށj
    
        'target-textbox��Left��source-textbox��Left+�C���f���g��
        Dim target_textbox_left As Double: target_textbox_left = .Left + indent_width
        'target-textbox��Top��source-textbox�̉��[
        Dim target_textbox_top As Double: target_textbox_top = .Top + .Height
        'target-textbox��Width��source-textbox�Ɠ���
        Dim target_textbox_width As Double: target_textbox_width = .width
        
        Set active_window = ActiveWindow.Selection.SlideRange '�����̓s����C���݃A�N�e�B�u�ȃX���C�h�̃I�u�W�F�N�g�𓾂Ă���
        
        'target-textbox�𐶐�
        Dim target_textbox As Shape: Set target_textbox = active_window.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=target_textbox_left, Top:=target_textbox_top, width:=target_textbox_width, Height:=target_textbox_height) 'AddTextbox(�����̌����C����_x���W�C����_y���W,�{�b�N�X�̕��C����)
        
        'Debug.Print tb_bold
        With target_textbox.TextFrame.TextRange
            .Text = target_text '.TextFrame.TextRange.Text ="���e"
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

'�|��
Function translate(source_text As String, source_lang As String, target_lang As String, api_key As String, exe_folder As String) As String

    ' Shell�֐����g�p����Python�X�N���v�g�����s
    Dim translate_cmd As String: translate_cmd = create_shell_arg(exe_folder & "\translator.exe") & create_shell_arg(source_text) & create_shell_arg(source_lang) & create_shell_arg(target_lang) & create_shell_arg(api_key) & create_shell_arg(exe_folder) ' xxx.exe ����1 ����2...�̌`���̃R�}���h
    Debug.Print translate_cmd '�R�}���h���e�m�F
    Shell translate_cmd, vbHide '�R�}���h����
    
    '�|��ꕶ���̃e�L�X�g�t�@�C�����i�[�����\��̃p�X
    Dim target_text_path As String
    target_text_path = exe_folder & "\translated.txt"
    'Debug.Print target_text_path
    
    '�|��ꕶ���̃e�L�X�g�t�@�C������������ĂȂ��Ȃ�΁C�܂��|��I����ĂȂ��̂ő҂�
    Do While Dir(target_text_path) = ""
        'Debug.Print "wait"
        Sleep 10 '���̃X���[�v���Ȃ��ƁC����while�����Ȃ��}���ɂ��G���[�N����
    Loop
        
    Dim target_text As String '�|��ꕶ��
    
    ' �|��ꕶ���̃e�L�X�g�t�@�C������ǂݍ���
    Open target_text_path For Input As #1
            Line Input #1, target_text
    Close #1
    
    Debug.Print target_text '�|��ꕶ���̊m�F
    'Debug.Print target_text_path
    Kill target_text_path ' �|��ꕶ���̃e�L�X�g�t�@�C���������@���̃t�@�C���̗L���Ŗ|��҂��𔻒f����̂ŁC�����Ă���
    
    translate = target_text

End Function


'���C���֐��݂����Ȃ���
Sub SlideByTransWriter()

    '===============================���͗�===========================================================================================================================================================
    
    '�|�󌾌�ݒ�l / Translation language setpoint / Example Japanese:JA, English:EN
    Dim source_lang As String: source_lang = "JA" '�|��O���� / pre-translated language
    Dim target_lang As String: target_lang = "EN" '�|��㌾�� / post-translational language
    
    '�|��T�[�r�X�܂��̐ݒ�l / Set values around translation services
    Dim api_key As String: api_key = "xxxxxxxxxxxxxx" 'DeepL API Key
    Dim exe_folder As String: exe_folder = "yyyyyyyyyyyyyyyyy" 'translator.exe�̂���t�H���_ / Folder with translator.exe
    
    'TextBox�̃t�H���g�̐ݒ�l / Font setting value for TextBox
    Dim tb_bold As Boolean: tb_bold = False '�����ɂ���Ȃ�True,�����łȂ��Ȃ�False
    Dim tb_color As Long: tb_color = RGB(166, 166, 166) '�t�H���g�̐F
    Dim tb_size As Double: tb_size = 20 '�t�H���g�T�C�Y
    Dim tb_fontname As String: tb_fontname = "Arial"
    
    'TextBox�̈ʒu��傫���̐ݒ�l / Set values for TextBox position and size
    Dim indent_width As Double: indent_width = 50 '�C���f���g��[pt] / indent width[pt]
    Dim target_textbox_height As Double: target_textbox_height = 40 '�|���e�L�X�g�{�b�N�X�̍���[pt]�D�t�H���g�T�C�Y���傫�����邱�ƁD/ Height of the text box after translation [pt]. Should be larger than the font size.
    
    '===============================���͗�===========================================================================================================================================================
        
    Dim source_text As String: source_text = get_select_shape_text() '���ݑI�����Ă���textbox�̃e�L�X�g=source-text���擾
    
    Dim target_text As String: target_text = translate(source_text, source_lang, target_lang, api_key, exe_folder) '�|�󂵁Ctarget-text���擾
    
    Call create_target_textbox(target_text, indent_width, target_textbox_height, tb_bold, tb_color, tb_size, tb_fontname) '���ݑI�����Ă���textbox�̉��ɂ���傫����target-textbox�𐶐�����
    
End Sub






