Private Sub Worksheet_Change(ByVal Target As Range)

  Application.EnableEvents = False

  If Range("C3") > 0 Then
    Range("C3").Offset(0, 2) = Range("C3") + Range("C3").Offset(0, 2)
    Range("C3").Value = 0
    If Range("G3").Value = 100 Then
      Range("I3").Value = "勇者"
    ElseIf Range("G3") >= 90 Then
      Range("I3").Value = "SSSランク冒険者"
    ElseIf Range("G3") >= 80 Then
      Range("I3").Value = "SSランク冒険者"
    ElseIf Range("G3") >= 70 Then
      Range("I3").Value = "Sランク冒険者"
    ElseIf Range("G3") >= 60 Then
      Range("I3").Value = "Aランク冒険者"
    ElseIf Range("G3") >= 50 Then
      Range("I3").Value = "Bランク冒険者"
    ElseIf Range("G3") >= 40 Then
      Range("I3").Value = "Cランク冒険者"
      Range("K3").Value = "レコーディングに挑戦してみよう！"
    ElseIf Range("G3") >= 30 Then
      Range("I3").Value = "Dランク冒険者"
      Range("K3").Value = "外側輪状披裂筋を鍛えて声の剛性感を高めよう！"
    ElseIf Range("G3") >= 20 Then
      Range("I3").Value = "Eランク冒険者"
      Range("K3").Value = "チェストをA5まで楽に出せるようにしよう！"
    ElseIf Range("G3") >= 10 Then
      Range("I3").Value = "Fランク冒険者"
      Range("K3").Value = "ファルセットをA6まで楽に出せるようにしよう！"
    Else
      Range("I3").Value = "村人"
    End If
  End If

  If Range("C4") > 0 Then
    Range("C4").Offset(0, 2) = Range("C4") + Range("C4").Offset(0, 2)
    Range("C4").Value = 0
    If Range("G3").Value = 100 Then
      Range("I3").Value = "勇者"
    ElseIf Range("G3") >= 90 Then
      Range("I3").Value = "SSSランク冒険者"
    ElseIf Range("G3") >= 80 Then
      Range("I3").Value = "SSランク冒険者"
    ElseIf Range("G3") >= 70 Then
      Range("I3").Value = "Sランク冒険者"
    ElseIf Range("G3") >= 60 Then
      Range("I3").Value = "Aランク冒険者"
    ElseIf Range("G3") >= 50 Then
      Range("I3").Value = "Bランク冒険者"
    ElseIf Range("G3") >= 40 Then
      Range("I3").Value = "Cランク冒険者"
    ElseIf Range("G3") >= 30 Then
      Range("I3").Value = "Dランク冒険者"
    ElseIf Range("G3") >= 20 Then
      Range("I3").Value = "Eランク冒険者"
    ElseIf Range("G3") >= 10 Then
      Range("I3").Value = "Fランク冒険者"
    Else
      Range("I3").Value = "村人"
      Range("K3").Value = "主要言語20種の制御構文をマスターしよう！"
    End If
  End If
  
  
  Application.EnableEvents = True

End Sub
