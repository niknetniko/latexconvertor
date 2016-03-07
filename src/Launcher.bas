Attribute VB_Name = "Launcher"
'''
' Start the convertor.
'''
Sub Auto_Open()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    Dim MyToolbar As String

    ' Give the toolbar a name
    MyToolbar = "Convertor"

    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)
    If Err.Number <> 0 Then
          ' The toolbar's already there, so we have nothing to do
          Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Now add a button to the new toolbar
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties

    With oButton

         .DescriptionText = "Convert"
          'Tooltip text when mouse if placed over button

         .Caption = "Begin converting"
         'Text if Text in Icon is chosen

         .OnAction = "Button1"
          'Runs the Sub Button1() code when clicked

         .Style = msoButtonIcon
          ' Button displays as icon, not text or both

         .FaceId = 527
          ' chooses icon #52 from the available Office icons

    End With

    ' Repeat the above for as many more buttons as you need to add
    ' Be sure to change the .OnAction property at least for each new button

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Visible = True

NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
     'Just in case there is an error
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Sub Button1()
' This code will run when you click Button 1 added above
' Add a similar subroutine for each additional button you create on the toolbar
    ' This is just some silly example code.
    ' You'd put your real working code here to do whatever
    ' it is that you want to do
    ProgressIndicator.Show
End Sub
