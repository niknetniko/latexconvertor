Attribute VB_Name = "LatexConvertor"
Option Private Module

'''
' Convert the powerpoint to LaTeX.
'''
Sub Convert_To_LaTeX()
    
    Dim sImagePath As String
    Dim tImagePath As String
    Dim sImageName As String
    Dim sPrefix As String
    Dim latexFile As String
    Dim counter As Integer
    Dim prevIndent As Integer
    Dim currIndent As Integer
    Dim total As Integer
    Dim slideCounter As Integer
    Dim Pgh As TextRange
    Dim oSlide As Slide '* Slide Object
    On Error GoTo Err_ImageSave

    sPrefix = Split(ActivePresentation.Name, ".")(0)
    sImagePath = ActivePresentation.path
    tImagePath = sImagePath & "\" & sPrefix
    
    ' If the directory does not exist, make it.
    If Len(Dir(tImagePath, vbDirectory)) = 0 Then
        MkDir tImagePath
    End If
    
    ' Beginning of the LaTeX file.
    latexFile = "\documentclass[11pt]{article}" & vbCrLf & "\usepackage{lmodern}" & vbCrLf & "\usepackage[T1]{fontenc}" & vbCrLf & "\usepackage[utf8]{inputenc}" & vbCrLf & "\usepackage{graphicx}" & vbCrLf & "\usepackage{a4wide}" & vbCrLf & "\begin{document}" & vbCrLf & "\setlength{\parskip}{\medskipamount}" & vbCrLf & "\setlength{\parindent}{0pt}" & vbCrLf
    
    total = ActivePresentation.Slides.Count
    slideCounter = 0
    For Each oSlide In ActivePresentation.Slides
        WriteMessage slideCounter, total
        sImageName = sPrefix & "-" & oSlide.SlideIndex & ".png"
        ' Export image with 1920x1440 (h x b), this needs to match the slide dimensions
        oSlide.Export tImagePath & "\" & sImageName, "PNG", 1920, 1440
        ' Add image to TeX file
        latexFile = latexFile & "\begin{center}" & vbCrLf & "\frame{\includegraphics[width=0.9\columnwidth]{" & sImageName & "}}" & vbCrLf & "\end{center}" & vbCrLf
        For Each oSh In oSlide.NotesPage.Shapes
            If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                If oSh.HasTextFrame Then
                    ' If the slide has notes
                    If Not oSh.TextFrame.TextRange Is Nothing Then
                        counter = 1
                        prevIndent = 1
                        ' Loop the paragraphs
                        For Each Pgh In oSh.TextFrame.TextRange.Paragraphs
                            currIndent = Pgh.Paragraphs.IndentLevel
                            
                            ' If the previous indent is lower than this one, begin item
                            If prevIndent < currIndent Then
                                latexFile = latexFile & "\begin{itemize}" & vbCrLf
                            ' If the previous indent is larger than this one, stop indent
                            ElseIf prevIndent > currIndent Then
                                latexFile = latexFile & "\end{itemize}" & vbCrLf
                            End If
                            
                            ' If it is the top paragraph, just add the text.
                            If currIndent = 1 Then
                                latexFile = latexFile & ConvertParagraph(Pgh)
                            ' Else add the text as subitem.
                            Else
                                latexFile = latexFile & "\item " + ConvertParagraph(Pgh)
                            End If
                            
                            ' Additional end itemize for the case when there is no more text after an itemize block.
                            If counter = oSh.TextFrame.TextRange.Paragraphs.Count And currIndent > 1 Then
                                latexFile = latexFile & "\end{itemize}" & vbCrLf
                            End If
                            
                            prevIndent = currIndent
                            counter = counter + 1
                        Next Pgh
                    End If
                End If
            End If
        Next oSh
        latexFile = latexFile & "\newpage{}" & vbCrLf & vbCrLf
        slideCounter = slideCounter + 1
    Next oSlide
    
    latexFile = latexFile & "\end{document}"
    
    WriteStringToFile tImagePath & "\" & sPrefix & ".tex", latexFile
    
    Done tImagePath & "\" & sPrefix & ".tex"

Err_ImageSave:
    If Err <> 0 Then
    MsgBox Err.Description
    End If
End Sub

'''
' Write a string to a file.
'
' @param String pFileName The name of the file. The path must exist.
' @param String pString The string to write to the file.
'
' @see http://stackoverflow.com/questions/31435662/vba-save-a-file-with-utf-8-without-bom/31436631#31436631 for BOM-related stuff
'''
Private Sub WriteStringToFile(fileName As String, toWrite As String)

    Const adSaveCreateNotExist = 1
    Const adSaveCreateOverWrite = 2
    Const adTypeBinary = 1
    Const adTypeText = 2
    
    Dim objStreamUTF8: Set objStreamUTF8 = CreateObject("ADODB.Stream")
    Dim objStreamUTF8NoBOM: Set objStreamUTF8NoBOM = CreateObject("ADODB.Stream")
    
    With objStreamUTF8
      .Charset = "UTF-8"
      .Open
      .WriteText toWrite
      .Position = 0
      .SaveToFile fileName, adSaveCreateOverWrite
      .Type = adTypeBinary
      .Position = 3
    End With
    
    With objStreamUTF8NoBOM
      .Type = adTypeBinary
      .Open
      objStreamUTF8.CopyTo objStreamUTF8NoBOM
      .SaveToFile fileName, adSaveCreateOverWrite
    End With

    'Dim fsT As Object
    'Set fsT = CreateObject("ADODB.Stream")
    'fsT.Type = 2 'Specify stream type - we want To save text/string data.
    'fsT.Charset = "utf-8" 'Specify charset For the source text data.
    'fsT.Open 'Open the stream And write binary data To the object
   ' fsT.WriteText toWrite
    'fsT.SaveToFile fileName, 2 'Save binary data To disk
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close

End Sub

'''
' Write a progress message to the form.
'
' @param Integer current The current progress.
' @param Integer total The total progress.
'''
Private Sub WriteMessage(current As Integer, total As Integer)

    Dim prct As Integer
    prct = CInt(Replace(Format(current / total, "0%"), "%", ""))
    ProgressIndicator.Text.Caption = current & " van " & total & " voltooid. (" & Format(current / total, "0%") & ")"
    ProgressIndicator.Bar.Width = prct * 2
    
    DoEvents
End Sub

'''
' Display that the converting was done.
'
' @param String path The path to the saved tex file.
'''
Private Sub Done(path As String)

    Dim prct As Integer
    prct = 100
    ProgressIndicator.Text.Caption = "Klaar."
    ProgressIndicator.Bar.Width = prct * 2
    ProgressIndicator.Done.Caption = "Opgeslagen als " & path
    ProgressIndicator.Cancel.Visible = False
    ProgressIndicator.DoneButton.Visible = True
    
    DoEvents
End Sub

'''
' Escape special LaTeX characters
'
' @return String The escaped character.
'''
Private Function ConvertCharForLatex(char As String) As String
    If char = "&" Or char = "%" Or char = "$" Or char = "#" Or char = "_" Or char = "{" Or char = "}" Or char = "~" Or char = "^" Or char = "\" Then
        char = "\" & char
    End If
    
    ConvertCharForLatex = char
End Function

'''
' Convert a paragraph to valid LaTeX-code. This accounts for sub- and superscript.
'
' @param TextRange paragraph The paragraph to convert.
'
' @return String The converted paragraph.
'''

Private Function ConvertParagraph(paragraph As TextRange) As String

    Dim superscript As String
    Dim subscript As String
    Dim superScripting As Boolean
    Dim subScripting As Boolean
    Dim result As String
    
    superScripting = False
    subScripting = False
    superscript = "\textsuperscript{"
    subscript = "\textsubscript{"
    result = ""
    
    For Each letter In paragraph.Characters
        If letter.Font.superscript Then
            superscript = superscript & ConvertCharForLatex(letter.Text)
            superScripting = True
        ElseIf superScripting Then
            result = result & superscript & "}" & ConvertCharForLatex(letter.Text)
            superscript = "\textsuperscript{"
            superScripting = False
        ElseIf letter.Font.subscript Then
            subscript = subscript & ConvertCharForLatex(letter.Text)
            subScripting = True
        ElseIf subScripting Then
            result = result & subscript & "}" & ConvertCharForLatex(letter.Text)
            subscript = "\textsubscript{"
            subScripting = False
        Else
            result = result & ConvertCharForLatex(letter.Text)
        End If
    Next letter
    
    ConvertParagraph = result
    
End Function
