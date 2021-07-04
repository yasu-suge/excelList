Attribute VB_Name = "Module2"
Option Explicit

'[VBA]OneDrive�œ������Ă���t�@�C���܂��̓t�H���_��URL�����[�J���p�X�ɕϊ�����֐�
'Copyright (c) 2020 �������̒� All Rights Reserved.
'This software is released under the GPLv3<https://opensource.org/licenses/GPL-3.0>.
'���̃\�t�g�E�F�A��GNU GPLv3�̉��Ń����[�X����Ă��܂�<https://opensource.org/licenses/GPL-3.0>�B;

'* @fn Public Function OneDriveUrlToLocalPath(ByRef Url As String) As String
'* @brief OneDrive�̃t�@�C��URL���̓t�H���_URL�����[�J���p�X�ɕϊ����܂��B
'* @param[in] Url OneDrive���ɕۑ����ꂽ�̃t�@�C�����̓t�H���_��URL
'* @return Variant ���[�J���p�X��Ԃ��܂��B����Url�Ƀ��[�J���p�X��"https://"�ȊO����n�܂镶������w�肵���ꍇ�A����Url��Ԃ��܂��B
'* @details OneDrive�̃t�@�C��URL���̓t�H���_URL�����[�J���p�X�ɕϊ����܂��B�{�֐��́AExcel�u�b�N��OneDrive���Ɋi�[����Ă���ꍇ�ɁAWorkbook.Path����Workbook.FullName��URL��Ԃ������������邽�߂̂��̂ł��B
'*
Public Function OneDriveUrlToLocalPath(ByRef Url As String) As String
Const OneDriveCommercialUrlPattern As String = "*my.sharepoint.com*" '�@�l����OneDrive��URL���ۂ��𔻒肷�邽�߂�Like�E�Ӓl

    '������URL�łȂ��ꍇ�A�����̓��[�J���p�X�Ɣ��f���Ă��̂܂ܕԂ��B
    If Not (Url Like "https://*") Then
        OneDriveUrlToLocalPath = Url
        Exit Function
    End If
    
    'OneDrive�̃p�X���擾���Ă���(�p�t�H�[�}���X�D��)�B
    Static PathSeparator As String
    Static OneDriveCommercialPath As String
    Static OneDriveConsumerPath As String
    
    If (PathSeparator = "") Then
        PathSeparator = Application.PathSeparator
        
        '�@�l����OneDrive(OneDrive for Business)�̃p�X
        OneDriveCommercialPath = Environ("OneDriveCommercial")
        If (OneDriveCommercialPath = "") Then OneDriveCommercialPath = Environ("OneDrive")
        
        '�l����OneDrive�̃p�X
        OneDriveConsumerPath = Environ("OneDriveConsumer")
        If (OneDriveConsumerPath = "") Then OneDriveConsumerPath = Environ("OneDrive")

    End If
    
    '�@�l����OneDrive�FURL��"https://��Ж�-my.sharepoint.com/personal/���[�U�[��_domain_com/Documents�t�@�C���p�X")
    Dim FilePathPos As Long
    If (Url Like OneDriveCommercialUrlPattern) Then
        FilePathPos = InStr(1, Url, "/Documents") + 10 '10 = Len("/Documents")
        OneDriveUrlToLocalPath = OneDriveCommercialPath & Replace(Mid(Url, FilePathPos), "/", PathSeparator)
        
    '�l����OneDrive�FURL��"https://d.docs.live.net/CID�ԍ�/�t�@�C���p�X"
    Else
        FilePathPos = InStr(9, Url, "/") '9 == Len("https://") + 1
        FilePathPos = InStr(FilePathPos + 1, Url, "/")

        If (FilePathPos = 0) Then
            OneDriveUrlToLocalPath = OneDriveConsumerPath
        Else
            OneDriveUrlToLocalPath = OneDriveConsumerPath & Replace(Mid(Url, FilePathPos), "/", PathSeparator)
        End If
    End If

End Function

