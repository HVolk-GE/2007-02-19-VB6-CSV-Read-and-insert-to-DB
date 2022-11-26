Attribute VB_Name = "OnlyCharaters"
Public Sub ChangeOnlyCharaters()

If sNewDecimal = sOldDecimal Or sNewSeparator = sOldSeparator Then
   MsgBox "Can't change the charaters, all same charaters have, please check these !", vbCritical
   Exit Sub
End If

If Form1.Check2.Value = 1 Then
    
    Form1.File1.Path = Form1.Text1.Text
    Form1.File1.FileName = Form1.Combo1.Text
    
    Form1.File2.Path = Form1.Text1.Text
    
    lZeilen0 = 0
    lZeilen3 = 0
    lZeilen2 = 0
    
    WriteZeile = ""
    WriteZeile0 = 0
    WriteZeile1 = ""
    WriteLine1 = ""
    
    j = 0
    i = Len(Form1.Combo1.Text)
 
    If i = 4 Then StrSoucreExte1 = Right(Form1.Combo1.Text, 3)
    If i = 5 Then StrSoucreExte1 = Right(Form1.Combo1.Text, 4)
 
    SourcePath = Form1.Text1.Text
    StrDesFile2 = Form1.Text2.Text
 
    StrSoucreFile0 = SourcePath & Form1.Combo2.Text & "." & ResFile

    StrSoucreFile1 = Form1.Combo2.Text & StrSoucreExte1
    StrSoucreFile2 = SourcePath & StrSoucreFile1
    StrSoucreFile3 = StrDesFile2 & AVdataFile
    StrSoucreFile4 = StrDesFile2 & ACdataFile
  
    Form1.ProgressBar1.Visible = True
    
    intProgress = Form1.File1.ListCount + 1
    
    intValueFiles = 100 / intProgress
    
    ' Hier ist es eine csv Datei:
    If i = 5 Then
         
         Form1.Command1.Enabled = False
         Form1.Command2.Enabled = False
         Form1.Combo1.Enabled = False
         Form1.Combo2.Enabled = False
         Form1.Text1.Enabled = False
         Form1.Text2.Enabled = False
         
        Form1.ProgressBar1.Value = intValueFiles * 1
                
        i = 0

       StrSoucreFile1 = SourcePath & StrSoucreFile1

        Open StrSoucreFile1 For Input As 1

        While Not EOF(1)
             ReDim Preserve sZeilen3(i) As String
             Line Input #1, sZeilen3(i)
             sZeilen3(i) = Replace(sZeilen3(i), sOldSeparator, sNewSeparator)
             sZeilen3(i) = Replace(sZeilen3(i), sOldDecimal, sNewDecimal)
             i = i + 1
        Wend
        
        Close #1
        
         i = i - 1
         
       StrSoucreFile1 = StrDesFile2 & Form1.Combo2.Text & StrSoucreExte1
         
        Open StrSoucreFile1 For Output As 1
        
        For a = 0 To i
            Print #1, sZeilen3(a)
        Next
        
        Close #1
        
    End If
        
         '#########################################################################
         '#
         '#  Jetzt wird ueberprueft ob es eine *.l* datei gibt, wenn ja wird auch
         '#  daraus die erste Spalte als ueberschrift gelesen und ...:
         '#
         '#########################################################################
          
     intValueFiles = 100 / intProgress
          
    If Form1.File2.ListCount > 0 Then
         
       For i = 0 To Form1.File2.ListCount - 1
            
          intVal = i + 1
            
          Form1.ProgressBar1.Value = intValueFiles * intVal / 2
          
          StrSoucreFile1 = SourcePath & Form1.File2.List(i)

          Open StrSoucreFile1 For Input As 1

          F = 0

          While Not EOF(1)
            ReDim Preserve sZeilen3(F) As String
            Line Input #1, sZeilen3(F)
            sZeilen3(F) = Replace(sZeilen3(F), sOldSeparator, sNewSeparator)
            sZeilen3(F) = Replace(sZeilen3(F), sOldDecimal, sNewDecimal)
            F = F + 1
          Wend
          
          F = F - 1
          
          Close #1
        
          StrSoucreFile1 = StrDesFile2 & Form1.File2.List(i)
          
          Open StrSoucreFile1 For Output As 1
          For a = 0 To F
            Print #1, sZeilen3(a)
          Next
        
          Close #1
       Next
                
    End If
End If
         Form1.Command2.Enabled = True
         Form1.Combo2.Enabled = True
         Form1.Text1.Enabled = True
         Form1.Text2.Enabled = True
         Form1.Combo2.Text = ""
         
         If Form1.Command1.Enabled = True Then
            Form1.Command1.Enabled = False
         End If
         
         If Form1.Check2.Value = 1 Then
            Form1.Check2.Value = 0
         End If

         Form1.ProgressBar1.Value = 0
         Form1.ProgressBar1.Visible = False
         
         MsgBox "Charaters changed, done !" & vbCr & _
         "For open these *.l??-Files change the extention to *.csv", vbInformation
End Sub
