Attribute VB_Name = "Module1"
Option Explicit

Dim rows As Long

Public Enum ColumnsPositon
    e_NO = 1
    e_WORK_BOOK_NAME
    e_DIR
    e_FULL_PATH
    e_WORK_SHEET_NAME
    e_IS_WORK_SHEET_VISIBLE
    e_IS_ALIGN_MARGINS_HEADER_FOTTER
    e_IS_BLACK_AND_WHITE
    e_TOP_MARGIN
    e_BOTTOM_MARGIN
    e_RIGHT_MARGIN
    e_LEFT_MARGIN
    e_HEADER_MARGIN
    e_FOOTER_MARGIN
    e_LEFT_HEADER
    e_CENTER_HEADER
    e_RIGHT_HEADER
    e_LEFT_FOOTER
    e_CENTER_FOOTER
    e_RIGHT_FOOTER
    e_PAPER_SIZE
    e_DRAFT
    e_PRINT_TITLE_COLUMNS
    e_PRINT_TITLE_ROWS
    e_PRINT_GRID_LINES
    e_ORIENTATION
    e_CENTER_HORIZONTALY
    e_CENTER_VERTICALLY
    e_FIRST_PAGE_NUMBER
    e_PRINT_AREA
    e_ZOOM
    e_PRINT_QUALITY
    e_PRINT_HEADINGS
    e_PRINT_COMMENT
    e_FIT_TO_PAGE_TALL
    e_FIT_TO_PAGE_WIDE
    e_AUTO_FILTER_MODE
    e_PROTECT_CONTENTS
    e_PROTECT_DRAWING_OBJECT
    e_WORKSHEET_TYPE
    e_ENABLE_CALCULATION
    e_WINDOW_COUNT
    e_WINDOW_ZOOM
    e_WINDOW_DISPLAY_MODE
    e_WINDOW_SPLIT
    e_WINDOW_FREEZE_PANES
    e_WINDOW_ACTIVE_CELL
End Enum

' 定数
Const XL_LANDSCAPE As String = "横向き"
Const XL_PORTAIT As String = "縦向き"
Const BOOL_EXIST As String = "あり"
Const BOOL_NOT_EXIST As String = "なし"
Const BOOL_YES As String = "はい"
Const BOOL_NO As String = "いいえ"
Const BOOL_DISPLAY As String = "表示"
Const BOOL_NON_DISPLAY As String = "非表示"
Const XL_AUTOMATIC As String = "自動"
Const XL_PRINT_IN_PLACE As String = "画面表示イメージ(メモのみ)"
Const XL_PRINT_NO_COMMENTS As String = "(なし)"
Const XL_PRINT_SHEET_END As String = "シートの末尾"
Const XL_CHART As String = "チャート"
Const XL_DIALOG_SHEET As String = "ダイアログ シート"
Const XL_EXCEL4_INTL_MACRO_SHEET As String = "Excel バージョン 4 International Macro シート"
Const XL_EXCEL4_MACRO_SHEET As String = "Excel バージョン 4 マクロ シート"
Const XL_WORKSHEET As String = "ワークシート"
Const XL_PAPER_10x14 As String = "10 in. x 14 in."
Const XL_PAPER_11x17 As String = "11 in. x 17 in."
Const XL_PAPER_A3 As String = "A3 (297 mm x 420 mm)"
Const XL_PAPER_A4 As String = "A4 (210 mm x 297 mm)"
Const XL_PAPER_A4_SMALL As String = "A4 Small (210 mm x 297 mm)"
Const XL_PAPER_A5 As String = "A5 (148 mm x 210 mm)"
Const XL_PAPER_B4 As String = "B4 (250 mm x 354 mm)"
Const XL_PAPER_B5 As String = "B5 (182 mm x 257 mm)"
Const XL_PAPER_CSHEET As String = "C サイズ シート"
Const XL_PAPER_DSHEET As String = "D サイズ シート"
Const XL_PAPER_ENVELOPE_10 As String = "封筒#10 (4-1/8 in. x 9-1/2 in.)"
Const XL_PAPER_ENVELOPE_11 As String = "封筒#11 (4-1/2 in. x 10-3/8 in.)"
Const XL_PAPER_ENVELOPE_12 As String = "封筒#12 (4-1/2 in. x 11 in.)"
Const XL_PAPER_ENVELOPE_14 As String = "封筒#14 (5 in. x 11-1/2 in.)"
Const XL_PAPER_ENVELOPE_9 As String = "封筒#9 (3-7/8 in. x 8-7/8 in.)"
Const XL_PAPER_ENVELOPE_B4 As String = "封筒 B4 (250 mm x 352 mm)"
Const XL_PAPER_ENVELOPE_B5 As String = "封筒 B5 (176 mm x 250 mm)"
Const XL_PAPER_ENVELOPE_B6 As String = "封筒 B6 (176 mm x 125 mm)"
Const XL_PAPER_ENVELOPE_C3 As String = "封筒 C3 (324 mm x 458 mm)"
Const XL_PAPER_ENVELOPE_C4 As String = "封筒 C4 (229 mm x 324 mm)"
Const XL_PAPER_ENVELOPE_C5 As String = "封筒 C5 (162 mm x 229 mm)"
Const XL_PAPER_ENVELOPE_C6 As String = "封筒 C6 (114 mm x 162 mm)"
Const XL_PAPER_ENVELOPE_C65 As String = "封筒 C65 (114 mm x 229 mm)"
Const XL_PAPER_ENVELOPE_DL As String = "封筒 DL (110 mm x 220 mm)"
Const XL_PAPER_ENVELOPE_ITALY As String = "封筒 (110 mm x 230 mm)"
Const XL_PAPER_ENVELOPE_MONARCH As String = "封筒モナーク(3-7/8 in. x 7-1/2 in.)"
Const XL_PAPER_ENVELOPE_PERSONAL As String = "封筒 (3-5/8 in. x 6-1/2 in.)"
Const XL_PAPER_ESHEET As String = "E サイズ シート"
Const XL_PAPER_EXECUTIVE As String = "エグゼクティブ (7- 1/2 in. x 10-1/2 in.)"
Const XL_PAPER_FANFOLD_LEGAL_GERMAN As String = "German Legal Fanfold(8-1/2 in. x 13 in.)"
Const XL_PAPER_FANFOLD_STD_GERMAN As String = "German Standard Fanfold(8-1/2 in. x 12 in.)"
Const XL_PAPER_FANFOLD_US As String = "U.S. Standard Fanfold(14-7/8 in.x 11 in.)"
Const XL_PAPER_FOLIO As String = "Folio (8-1/2 in. x 13 in.)"
Const XL_PAPER_LEDGER As String = "台帳 (17 in. x 11 in.)"
Const XL_PAPER_LEGAL As String = "Legal (8-1/2 in. x 14 in.)"
Const XL_PAPER_LETTER As String = "レター (8-1/2 in. x 11 in.)"
Const XL_PAPER_LETTER_SMALL As String = "レター Small (8-1/2 in. x 11 in.)"
Const XL_PAPER_NOTE As String = "ノート (8-1/2 in. x 11 in.)"
Const XL_PAPER_QUARTO As String = "4 つ折版 (215 mm x 275 mm)"
Const XL_PAPER_STATEMENT As String = "ステートメント (5- 1/2 in. x 8-1/2 in.)"
Const XL_PAPER_TABLOID As String = "タブロイド (11 in. x 17 in.)"
Const XL_PAPER_USER As String = "ユーザー定義"
Const XL_NORMAL_VIEW As String = "標準"
Const XL_PAGE_BREAK_PREVIEW As String = "改ページプレビュー"
Const XL_PAGE_LAYOUT_VIEW As String = "ページ レイアウト ビュー"

' メイン処理
Sub createList()
    Dim targetFolder As String
    Dim fso As Object
    
    ' 画面更新の停止
    Application.ScreenUpdating = False
    'イベント抑止
    Application.EnableEvents = False
    
    
    ' 対象フォルダの指定
    targetFolder = ThisWorkbook.Worksheets(1).Range("B1").value
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 対象フォルダ配下（サブフォルダ）の全ファイルに対する処理（再帰処理）
    Call loopAllFiles(targetFolder, fso)
    
    Set fso = Nothing
    
    'イベント抑止を解除
    Application.EnableEvents = True
    ' 画面更新の停止
    Application.ScreenUpdating = True
    MsgBox prompt:="処理が終了しました。"

End Sub

' 対象フォルダ配下（サブフォルダ）の全ファイルに対する処理（再帰処理）
Private Function loopAllFiles(targetFolder As String, fso As Object)

    Const FILE_TYPE_XLSX As String = "xlsx"
    Const FILE_TYPE_XLS As String = "xls"
    
    Dim folder As Object
    Dim file As Object
    
    rows = 2
    
    'サブフォルダの数だけ再起
    For Each folder In fso.getFolder(targetFolder).SubFolders
        Call loopAllFiles(folder.PATH, fso)
    Next folder
    
    'ファイルの数分繰り返し
    For Each file In fso.getFolder(targetFolder).Files
    
        Dim extentionName As String
        extentionName = fso.GetExtensionName(file.name)
        
        If LCase(extentionName) = FILE_TYPE_XLS Or LCase(extentionName) = FILE_TYPE_XLSX Then
            ' Excelファイルに対する処理
            Call execExcelFile(file)
        End If
    
    Next file

End Function

' Excelファイルに対する処理
Private Function execExcelFile(file As Object)

    Dim wkbook As Workbook
    
    Set wkbook = Workbooks.Open(Filename:=file.PATH, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, ReadOnly:=True)

    Debug.Print "ワークブック名" + wkbook.name

    Dim wksheet As Worksheet
    For Each wksheet In wkbook.Worksheets
        ' ワークシートに対する処理
        Call execWorksheet(wksheet, wkbook)
        rows = rows + 1
    Next wksheet
    
    wkbook.Close SaveChanges:=False

End Function

' ワークシートに対する処理
Private Function execWorksheet(wksheet As Worksheet, wkbook As Workbook)

    Debug.Print "ワークシート名：" + wksheet.name
    
    Dim listWkSheet As Worksheet
    Set listWkSheet = ThisWorkbook.Worksheets("リスト")
    
    listWkSheet.Cells(rows, ColumnsPositon.e_NO).value = rows - 1
    listWkSheet.Cells(rows, ColumnsPositon.e_WORK_BOOK_NAME).value = getWorkBookName(wkbook)
    listWkSheet.Cells(rows, ColumnsPositon.e_DIR).value = OneDriveUrlToLocalPath(getPath(wkbook))
    listWkSheet.Cells(rows, ColumnsPositon.e_FULL_PATH).value = OneDriveUrlToLocalPath(getFullPath(wkbook))
    listWkSheet.Cells(rows, ColumnsPositon.e_WORK_SHEET_NAME).value = getWorkSheetName(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_IS_WORK_SHEET_VISIBLE).value = isVisibleWorkSheet(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_IS_ALIGN_MARGINS_HEADER_FOTTER).value = isAlignMarginsHeaderFooter(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_IS_BLACK_AND_WHITE).value = isBlackAndWhite(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_TOP_MARGIN).value = getTopMargin(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_BOTTOM_MARGIN).value = getBottomMargin(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_LEFT_MARGIN).value = getLeftMargin(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_RIGHT_MARGIN).value = getRightMargin(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_HEADER_MARGIN).value = getHeaderMargin(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_FOOTER_MARGIN).value = getFooterMargin(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_LEFT_HEADER).value = getLeftHeader(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_CENTER_HEADER).value = getCenterHeader(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_RIGHT_HEADER).value = getRightHeader(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_LEFT_FOOTER).value = getLeftFooter(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_CENTER_FOOTER).value = getCenterFooter(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_RIGHT_FOOTER).value = getRightFooter(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PAPER_SIZE).value = getPaperSize(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_DRAFT).value = isDraft(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_TITLE_COLUMNS).value = getPrintTitleColumns(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_TITLE_ROWS).value = getPrintTitleRows(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_GRID_LINES).value = isPrintGridlines(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_ORIENTATION).value = getOrientation(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_CENTER_HORIZONTALY).value = isCenterHorizontally(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_CENTER_VERTICALLY).value = isCenterVertically(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_FIRST_PAGE_NUMBER).value = getFirstPageNumber(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_AREA).value = getPrintArea(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_ZOOM).value = getZoom(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_QUALITY).value = getPrintQuality(wksheet, 1)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_HEADINGS).value = isPrintHeadings(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PRINT_COMMENT).value = isPrintComments(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_FIT_TO_PAGE_TALL).value = getFitToPagesTall(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_FIT_TO_PAGE_WIDE).value = getFitToPagesWide(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_AUTO_FILTER_MODE).value = isAutoFilterMode(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PROTECT_CONTENTS).value = isProtectContens(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_PROTECT_DRAWING_OBJECT).value = isProtectDrawingObjects(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_WORKSHEET_TYPE).value = getType(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_ENABLE_CALCULATION).value = isEnableCalculation(wksheet)
    listWkSheet.Cells(rows, ColumnsPositon.e_WINDOW_COUNT).value = getWindowCount(wkbook)
    listWkSheet.Cells(rows, ColumnsPositon.e_WINDOW_ZOOM).value = getWindowZoom(wkbook)
    listWkSheet.Cells(rows, ColumnsPositon.e_WINDOW_DISPLAY_MODE).value = getWindowDisplayMode(wkbook)
    listWkSheet.Cells(rows, ColumnsPositon.e_WINDOW_SPLIT).value = isSplit(wkbook)
    listWkSheet.Cells(rows, ColumnsPositon.e_WINDOW_FREEZE_PANES).value = isFreezePain(wkbook)
    listWkSheet.Cells(rows, ColumnsPositon.e_WINDOW_ACTIVE_CELL).value = getActiveCell(wkbook)
    
    
    
End Function

' ワークブック名の取得
Private Function getWorkBookName(wkbook As Workbook) As String
    getWorkBookName = wkbook.name
End Function

' ディレクトリの取得
Private Function getPath(wkbook As Workbook) As String
    getPath = wkbook.PATH
End Function

' フルパスの取得
Private Function getFullPath(wkbook As Workbook) As String
    getFullPath = wkbook.FullName
End Function

' ワークシート名
Private Function getWorkSheetName(wksheet As Worksheet) As String
    getWorkSheetName = wksheet.name
End Function

' ワークシート表示非表示
Private Function isVisibleWorkSheet(wksheet As Worksheet) As String
    
    Dim bool As Boolean
    bool = wksheet.Visible
    
    Dim result As String
    If bool = True Then
        result = BOOL_DISPLAY
    Else
        result = BOOL_NON_DISPLAY
    End If
    
    isVisibleWorkSheet = result
    
End Function

' はい、いいえ
Private Function displayYesNo(bool As Boolean) As String
    Dim result As String

    If bool = True Then
        result = BOOL_YES
    Else
        result = BOOL_NO
    End If
    
    displayYesNo = result

End Function

' はい、いいえ
Private Function displayYesNo2(value As String) As Boolean
    Dim result As Boolean
    Select Case value
        Case BOOL_YES
            result = True
        Case BOOL_NO
            result = False
    End Select
    displayYesNo2 = result
End Function

' ページの余白に合わせて配置
Private Function isAlignMarginsHeaderFooter(wksheet As Worksheet) As String
    isAlignMarginsHeaderFooter = displayYesNo(wksheet.PageSetup.AlignMarginsHeaderFooter)
End Function

' 白黒印刷
Private Function isBlackAndWhite(wksheet As Worksheet) As String
    isBlackAndWhite = displayYesNo(wksheet.PageSetup.BlackAndWhite)
End Function

' 上余白(センチメートル)
Private Function getTopMargin(wksheet As Worksheet) As Double
    getTopMargin = Round(wksheet.PageSetup.TopMargin / Application.CentimetersToPoints(1#), 2)
End Function

' 下余白(センチメートル)
Private Function getBottomMargin(wksheet As Worksheet) As Double
    getBottomMargin = Round(wksheet.PageSetup.BottomMargin / Application.CentimetersToPoints(1#), 2)
End Function

' 右余白(センチメートル)
Private Function getRightMargin(wksheet As Worksheet) As Double
    getRightMargin = Round(wksheet.PageSetup.RightMargin / Application.CentimetersToPoints(1#), 2)
End Function

' 左余白(センチメートル)
Private Function getLeftMargin(wksheet As Worksheet) As Double
    getLeftMargin = Round(wksheet.PageSetup.LeftMargin / Application.CentimetersToPoints(1#), 2)
End Function

' ヘッダマージン(センチメートル)
Private Function getHeaderMargin(wksheet As Worksheet) As Double
    getHeaderMargin = Round(wksheet.PageSetup.HeaderMargin / Application.CentimetersToPoints(1#), 2)
End Function

' フッタマージン(センチメートル)
Private Function getFooterMargin(wksheet As Worksheet) As Double
    getFooterMargin = Round(wksheet.PageSetup.FooterMargin / Application.CentimetersToPoints(1#), 2)
End Function

' 左ヘッダ
Private Function getLeftHeader(wksheet As Worksheet) As String
    getLeftHeader = wksheet.PageSetup.LeftHeader
End Function

' 中央ヘッダ
Private Function getCenterHeader(wksheet As Worksheet) As String
    getCenterHeader = wksheet.PageSetup.CenterHeader
End Function

' 右ヘッダ
Private Function getRightHeader(wksheet As Worksheet) As String
    getRightHeader = wksheet.PageSetup.RightHeader
End Function

' 左フッタ
Private Function getLeftFooter(wksheet As Worksheet) As Variant
    getLeftFooter = wksheet.PageSetup.LeftFooter
End Function

' 中央フッタ
Private Function getCenterFooter(wksheet As Worksheet) As String
    getCenterFooter = wksheet.PageSetup.CenterFooter
End Function

' 右フッタ
Private Function getRightFooter(wksheet As Worksheet) As String
    getRightFooter = wksheet.PageSetup.RightFooter
End Function

' 用紙サイズ
Private Function getPaperSize(wksheet As Worksheet) As String
    Dim paperSizeName As Object
    Set paperSizeName = CreateObject("Scripting.Dictionary")
    
    paperSizeName.Add XlPaperSize.xlPaper10x14, XL_PAPER_10x14
    paperSizeName.Add XlPaperSize.xlPaper11x17, XL_PAPER_11x17
    paperSizeName.Add XlPaperSize.xlPaperA3, XL_PAPER_A3
    paperSizeName.Add XlPaperSize.xlPaperA4, XL_PAPER_A4
    paperSizeName.Add XlPaperSize.xlPaperA4Small, XL_PAPER_A4_SMALL
    paperSizeName.Add XlPaperSize.xlPaperA5, XL_PAPER_A5
    paperSizeName.Add XlPaperSize.xlPaperB4, XL_PAPER_B4
    paperSizeName.Add XlPaperSize.xlPaperB5, XL_PAPER_B5
    paperSizeName.Add XlPaperSize.xlPaperCsheet, XL_PAPER_CSHEET
    paperSizeName.Add XlPaperSize.xlPaperDsheet, XL_PAPER_DSHEET
    paperSizeName.Add XlPaperSize.xlPaperEnvelope10, XL_PAPER_ENVELOPE_10
    paperSizeName.Add XlPaperSize.xlPaperEnvelope11, XL_PAPER_ENVELOPE_11
    paperSizeName.Add XlPaperSize.xlPaperEnvelope12, XL_PAPER_ENVELOPE_12
    paperSizeName.Add XlPaperSize.xlPaperEnvelope14, XL_PAPER_ENVELOPE_14
    paperSizeName.Add XlPaperSize.xlPaperEnvelope9, XL_PAPER_ENVELOPE_9
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeB4, XL_PAPER_ENVELOPE_B4
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeB5, XL_PAPER_ENVELOPE_B5
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeB6, XL_PAPER_ENVELOPE_B6
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeC3, XL_PAPER_ENVELOPE_C3
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeC4, XL_PAPER_ENVELOPE_C4
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeC5, XL_PAPER_ENVELOPE_C5
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeC6, XL_PAPER_ENVELOPE_C6
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeC65, XL_PAPER_ENVELOPE_C65
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeDL, XL_PAPER_ENVELOPE_DL
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeItaly, XL_PAPER_ENVELOPE_ITALY
    paperSizeName.Add XlPaperSize.xlPaperEnvelopeMonarch, XL_PAPER_ENVELOPE_MONARCH
    paperSizeName.Add XlPaperSize.xlPaperEnvelopePersonal, XL_PAPER_ENVELOPE_PERSONAL
    paperSizeName.Add XlPaperSize.xlPaperEsheet, XL_PAPER_ESHEET
    paperSizeName.Add XlPaperSize.xlPaperExecutive, XL_PAPER_EXECUTIVE
    paperSizeName.Add XlPaperSize.xlPaperFanfoldLegalGerman, XL_PAPER_FANFOLD_LEGAL_GERMAN
    paperSizeName.Add XlPaperSize.xlPaperFanfoldStdGerman, XL_PAPER_FANFOLD_STD_GERMAN
    paperSizeName.Add XlPaperSize.xlPaperFanfoldUS, XL_PAPER_FANFOLD_US
    paperSizeName.Add XlPaperSize.xlPaperFolio, XL_PAPER_FOLIO
    paperSizeName.Add XlPaperSize.xlPaperLedger, XL_PAPER_LEDGER
    paperSizeName.Add XlPaperSize.xlPaperLegal, XL_PAPER_LEGAL
    paperSizeName.Add XlPaperSize.xlPaperLetter, XL_PAPER_LETTER
    paperSizeName.Add XlPaperSize.xlPaperLetterSmall, XL_PAPER_LETTER_SMALL
    paperSizeName.Add XlPaperSize.xlPaperNote, XL_PAPER_NOTE
    paperSizeName.Add XlPaperSize.xlPaperQuarto, XL_PAPER_QUARTO
    paperSizeName.Add XlPaperSize.xlPaperStatement, XL_PAPER_STATEMENT
    paperSizeName.Add XlPaperSize.xlPaperTabloid, XL_PAPER_TABLOID
    paperSizeName.Add XlPaperSize.xlPaperUser, XL_PAPER_USER
    
    Dim result As String
    result = paperSizeName.Item(wksheet.PageSetup.PaperSize)
    Set paperSizeName = Nothing
    
    getPaperSize = result
End Function

' ドラフト印刷
Private Function isDraft(wksheet As Worksheet) As String
    isDraft = displayYesNo(wksheet.PageSetup.DRAFT)
End Function

' 印刷タイトル行
Private Function getPrintTitleColumns(wksheet As Worksheet) As String
    getPrintTitleColumns = wksheet.PageSetup.PrintTitleColumns
End Function

' 印刷タイトル列
Private Function getPrintTitleRows(wksheet As Worksheet) As String
    getPrintTitleRows = wksheet.PageSetup.PrintTitleRows
End Function

' あり、なし
Private Function displayExist(bool As Boolean) As String
    Dim result As String
    If bool = True Then
        result = BOOL_EXIST
    Else
        result = BOOL_NOT_EXIST
    End If
    
    displayExist = result
End Function

' セルの枠線
Private Function isPrintGridlines(wksheet As Worksheet) As String
    isPrintGridlines = displayExist(wksheet.PageSetup.PrintGridlines)
End Function

' 印刷の向き
Private Function getOrientation(wksheet As Worksheet) As String
    
    Dim pageOrientationName As Object
    Set pageOrientationName = CreateObject("Scripting.Dictionary")
    
    pageOrientationName.Add XlPageOrientation.xlLandscape, XL_LANDSCAPE
    pageOrientationName.Add XlPageOrientation.xlPortrait, XL_PORTAIT
    
    Dim result As String
    result = pageOrientationName.Item(wksheet.PageSetup.ORIENTATION)
    Set pageOrientationName = Nothing
    
    getOrientation = result
End Function

' 印刷の向き
Private Function getOrientation2(value As String) As XlPageOrientation
    Dim result As XlPageOrientation
    Select Case value
    Case XL_LANDSCAPE
        result = XlPageOrientation.xlLandscape
    Case XL_PORTAIT
        result = XlPageOrientation.xlPortrait
    End Select
    getOrientation = result
End Function

' ページ横中央に配置
Private Function isCenterHorizontally(wksheet As Worksheet) As String
    isCenterHorizontally = displayYesNo(wksheet.PageSetup.CenterHorizontally)
End Function

' ページ縦中央に配置
Private Function isCenterVertically(wksheet As Worksheet) As String
    isCenterVertically = displayYesNo(wksheet.PageSetup.CenterVertically)
End Function

' 先頭ページ番号
Private Function getFirstPageNumber(wksheet As Worksheet) As String
    Dim firstPageNumber As Long
    Dim result As String
    
    firstPageNumber = wksheet.PageSetup.firstPageNumber
    
    If firstPageNumber = xlAutomatic Then
        result = "自動"
    Else
        result = firstPageNumber
    End If
    getFirstPageNumber = result
End Function

' 先頭ページ番号
Private Function getFirstPageNumber2(value As String) As Long
    Dim result As Long
    If value = XL_AUTOMATIC Then
        result = xlAutomatic
    Else
        result = value
    End If
    getFirstPageNumber2 = result

End Function

' 印刷範囲
Private Function getPrintArea(wksheet As Worksheet) As String
    getPrintArea = wksheet.PageSetup.PrintArea
End Function

' 拡大縮小率
Private Function getZoom(wksheet As Worksheet) As String
    Dim zoomValue As Variant
    zoomValue = wksheet.PageSetup.Zoom
    Dim result As String
    
    If zoomValue = False Then
        result = XL_AUTOMATIC
    Else
        result = zoomValue
    End If
    getZoom = result
End Function

' 印刷品質
Private Function getPrintQuality(wksheet As Worksheet, index As Integer) As String
    getPrintQuality = wksheet.PageSetup.PrintQuality(index)
End Function

' 行見出し・列見出し
Private Function isPrintHeadings(wksheet As Worksheet) As String
    isPrintHeadings = displayYesNo(wksheet.PageSetup.PrintHeadings)
End Function

' コメント印刷
Private Function isPrintComments(wksheet As Worksheet) As String
    Dim commentLocation As XlPrintLocation
    commentLocation = wksheet.PageSetup.PrintComments
    Dim result As String
    
    Select Case commentLocation
        Case XlPrintLocation.xlPrintInPlace
            result = XL_PRINT_IN_PLACE
        Case XlPrintLocation.xlPrintNoComments
            result = XL_PRINT_NO_COMMENTS
        Case XlPrintLocation.xlPrintSheetEnd
            result = XL_PRINT_SHEET_END
    End Select
    
    isPrintComments = result
    
End Function

' 拡大縮小ページ高さの数
Private Function getFitToPagesTall(wksheet As Worksheet) As String
    getFitToPagesTall = wksheet.PageSetup.FitToPagesTall
End Function

' 拡大縮小する幅のページ数
Private Function getFitToPagesWide(wksheet As Worksheet) As String
    getFitToPagesWide = wksheet.PageSetup.FitToPagesWide
End Function

' オートフィルタモード
Private Function isAutoFilterMode(wksheet As Worksheet) As String
    isAutoFilterMode = displayYesNo(wksheet.AutoFilterMode)
End Function

' シートの保護
Private Function isProtectContens(wksheet As Worksheet) As String
    isProtectContens = displayYesNo(wksheet.ProtectContents)
End Function

' 図形の保護
Private Function isProtectDrawingObjects(wksheet As Worksheet) As String
    isProtectDrawingObjects = displayYesNo(wksheet.ProtectDrawingObjects)
End Function

' ワークシートの種類
Private Function getType(wksheet As Worksheet) As String
    Dim sheetType As XlSheetType
    sheetType = wksheet.Type
    Dim result As String
    
    Select Case sheetType
        Case XlSheetType.xlChart
            result = XL_CHART
        Case XlSheetType.xlDialogSheet
            result = XL_DIALOG_SHEET
        Case XlSheetType.xlExcel4IntlMacroSheet
            result = XL_EXCEL4_INTL_MACRO_SHEET
        Case XlSheetType.xlExcel4MacroSheet
            result = XL_EXCEL4_MACRO_SHEET
        Case XlSheetType.xlWorksheet
            result = XL_WORKSHEET
    End Select
        
    getType = result
End Function

' 再計算の有無
Private Function isEnableCalculation(wksheet As Worksheet) As String
    isEnableCalculation = displayExist(wksheet.EnableCalculation)
End Function

' ウィンドウの数
Private Function getWindowCount(wkbook As Workbook) As Long
    getWindowCount = wkbook.windows.Count
End Function

' ウィンドウ表示倍率
Private Function getWindowZoom(wkbook As Workbook) As Variant
    getWindowZoom = wkbook.windows(1).Zoom
End Function

' ウィンドウ表示モード
Private Function getWindowDisplayMode(wkbook As Workbook) As String
    Dim view As XlWindowView
    view = wkbook.windows(1).view
    Dim result As String
    If view = xlNormalView Then
        result = XL_NORMAL_VIEW
    ElseIf view = xlPageBreakPreview Then
        result = XL_PAGE_BREAK_PREVIEW
    Else
        result = XL_PAGE_LAYOUT_VIEW
    End If
    
    getWindowDisplayMode = result
    
End Function

' ウィンドウの分割
Private Function isSplit(wkbook As Workbook) As String
    isSplit = displayYesNo(wkbook.windows(1).Split)
End Function


' 分割されたウィンドウが固定
Private Function isFreezePain(wkbook As Workbook) As String
    isFreezePain = displayYesNo(wkbook.windows(1).FreezePanes)
End Function

' アクティブセル
Private Function getActiveCell(wkbook As Workbook) As String
    Dim wnd As window
    wkbook.windows(1).Activate
    Dim cell As Range
    cell = ActiveCell
    Dim aaa As String
    aaa = cell.Address(ReferenceStyle:=xlR1C1)
    getActiveCell = cell
    
End Function

