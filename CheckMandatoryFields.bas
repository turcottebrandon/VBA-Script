Public Function CheckMandatoryFields(parentForm As String, Optional searchCriteria As String = "M") As Boolean
'Search 'Tag' field on certain controls on a specified form (Textbox, Combobox, Option group) for a criteria.
'If the criteria is not met (e.g. the field is blank or <0) the field will be highlighted in red and a message displayed.
'
'If any criteria are not met, the function returns 'False'
'
'Multiple characters can be entered into the 'Tag' field as the function uses InStr() function to look for matches.
'Example:  Control-1 Tag set = 'ABC', Control-2 Tag set = 'C'
'  Tst = CheckMandatoryFields("frmFoo","A")   -> would only check criteria against Control-1
'  Tst = CheckMandatoryFields("frmFoo","BC")   -> would only check criteria against Control-1
'  Tst = CheckMandatoryFields("frmFoo","C")   -> would check criteria against Control-1 and Control-2

'''''INITIALIZE
1   Const METHOD_NAME As String = "CheckMandatoryFields"  'for use with Error handling
5   On Error GoTo errHandler:

''''DECLARATIONS
  Dim frm As Form
  Dim MandatoryCount As Integer
  Dim ctl as Control

''''EXECUTE
30    Set frm = Forms(parentForm).[NavigationSubform].Form
35      Dim ctl As Control
40      For Each ctl In frm.Controls
45        If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acOptionGroup) Then
59          If InStr(1, ctl.Properties("Tag"), searchCriteria) > 0 And ctl.Properties("Visible") = True Then 'found a mandatory field
55              If IsNull(ctl.value) Or ctl.value = 0 Or ctl.value = "" Then 'if criteria not met, color the border red
60                  ctl.Properties("BorderColor") = vbRed
65                  MandatoryCount = MandatoryCount + 1
                Else
70                  ctl.Properties("BorderColor") = vbBlack 'if criteria met, color the border black
75                  MandatoryCount = MandatoryCount + 0
                End If
            End If
95        End If
100     Next ctl

105     If MandatoryCount = 0 Then
110        CheckMandatoryFields = True  'returns TRUE is all mandatory fields are not empty/null
115     Else
120         CheckMandatoryFields = False 'returns false if a field does not meet criteria
125         a = MsgBox("Sorry, you need to complete all mandatory fields", vbCritical, "Missing Information")
130     End If



'''''ERROR HANDLING
errHandler:
    If Err.Number > 0 Then
        Call errHandler(Err.Number, Erl, METHOD_NAME)   'common error handling routine which returns Method and Line number
    End If

End Function 
