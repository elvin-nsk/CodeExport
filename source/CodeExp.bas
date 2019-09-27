Attribute VB_Name = "CodeExp"
Option Explicit

Public Sub Settings()
  form.Show
End Sub

Public Sub Export()
  form.ExportQueue
  Unload form
End Sub
