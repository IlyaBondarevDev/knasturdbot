START_MESSAGE = '''
Привет, {username}! 
Отправь мне файл со своей работой и я его исправлю.
'''

DOWNLOADED_FILE_MESSAGE = '''
Я не могу обработать файл {filename}.
'''

DESTINATION_FILES = '''
f'files/{chat_id}/{file_name}'
'''

SUB_NAME = 'FormattingDocument'

MACROS_CODE = f'''
Sub ChangeFontAndSpacing()
    Selection.WholeStory
    
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 14
    End With
    
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub

Sub ChangeMargins()
    Selection.WholeStory

    With ActiveDocument.PageSetup
        .TopMargin = MillimetersToPoints(5)
        .BottomMargin = MillimetersToPoints(5)
        .LeftMargin = MillimetersToPoints(20)
        .RightMargin = MillimetersToPoints(5)
    End With
    
End Sub

Sub AddPageNumbers()
    Selection.WholeStory
    ActiveDocument.PageSetup.FooterDistance = CentimetersToPoints(1.25)
    
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers
        .Add PageNumberAlignment:=wdAlignPageNumberCenter
        .NumberStyle = wdPageNumberStyleArabic
        .HeadingLevelForChapter = 0
        .IncludeChapterNumber = False
        .ChapterPageSeparator = wdSeparatorHyphen
        .RestartNumberingAtSection = False
        .StartingNumber = 0
    End With
    
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Font
        .Name = "Times New Roman"
        .Size = 12
    End With
    
    ActiveDocument.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
End Sub

Sub AlignPictures()
    Dim picNum As Integer
    Dim picName As String
    Dim rng As Range
    
    picNum = 1
    
    For Each pic In ActiveDocument.InlineShapes
        pic.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

        pic.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        pic.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        pic.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        pic.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        
        Set rng = pic.Range
        rng.Collapse wdCollapseEnd
        rng.InsertParagraphAfter
        rng.Collapse wdCollapseEnd

        picName = "Название Рисунка"
        
        rng.Text = "Рисунок " & picNum & " – " & picName
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
        picNum = picNum + 1
    Next pic
End Sub

Sub FindBoldAndApplyHeading()
    Dim myStyle As Style
    Set myStyle = ActiveDocument.Styles.Add(Name:="Мой стиль", Type:=wdStyleTypeParagraph)
    With myStyle
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .ParagraphFormat.FirstLineIndent = InchesToPoints(0.5)
        .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
        .ParagraphFormat.OutlineLevel = wdOutlineLevel1
    End With
    
    Dim rngSearch As Range
    Set rngSearch = ActiveDocument.Content
    Do While True
        With rngSearch.Find
            .Font.Bold = True
            .Execute
        End With
        If rngSearch.Find.Found Then
            rngSearch.Style = myStyle
        Else
            Exit Do
        End If
    Loop
End Sub

Sub AddTableOfContentsOnThirdPage()
    Dim rngInsert As Range
    Set rngInsert = ActiveDocument.GoTo(wdGoToPage, wdGoToAbsolute, 3)
    rngInsert.InsertBreak Type:=wdSectionBreakNextPage
    ActiveDocument.TablesOfContents.Add Range:=rngInsert, UseHeadingStyles:=True, _
        UpperHeadingLevel:=1, LowerHeadingLevel:=3, IncludePageNumbers:=True, _
        RightAlignPageNumbers:=True, UseHyperlinks:=True
End Sub

Sub ${SUB_NAME}()
    AddPageNumbers
    ChangeMargins
    AlignPictures
    FindBoldAndApplyHeading
    AddTableOfContentsOnThirdPage
End Sub
'''