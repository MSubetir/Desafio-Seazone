Attribute VB_Name = "Módulo1"
Sub Ss():

Dim lastCell As Range

Set lastCell = Planilha1.Range("B1000").End(xlUp)

plot = "IT2"
monthStart = "HH"
monthEnd = "IL"


Planilha1.Range(plot).Select

Do While ActiveCell.Row <= lastCell.Row

Set intervalSom = Range(monthStart & ActiveCell.Row & ":" & monthEnd & ActiveCell.Row)
soma = 0

For Each cel In intervalSom
    If cel.Interior.ColorIndex = 4 Then
        soma = soma + cel.Value
    End If
Next

ActiveCell.Value = soma
ActiveCell.Offset(1, 0).Select

Loop


End Sub
