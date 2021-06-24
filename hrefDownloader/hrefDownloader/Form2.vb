Public Class Form2
    Public CData As ChemData

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim o As String = ""
        My.Computer.FileSystem.WriteAllText(TextBox1.Text, o, False)
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            o &= CheckedListBox1.Items.Item(i) & ","
        Next
        o = Mid(o, 1, o.Length - 1) & vbCrLf
        Dim dig As Integer = NumLib.Digitof(CData.Species.Count - 1)
        For i As Integer = 0 To CData.Species.Count - 1
            If CheckedListBox1.GetItemChecked(0) Then
                o &= ToCSVItemFormat(NumLib.NumFormat(i, dig)) & ","
            End If
            If CheckedListBox1.GetItemChecked(1) Then
                o &= ToCSVItemFormat(CData.Species(i).Name) & ","
            End If
            If CheckedListBox1.GetItemChecked(2) Then
                o &= ToCSVItemFormat(CData.Species(i).ReactionFormula) & ","
            End If
            If CheckedListBox1.GetItemChecked(3) Then
                o &= ToCSVItemFormat(CData.Species(i).Formula) & ","
            End If
            If CheckedListBox1.GetItemChecked(4) Then
                Dim t As Double
                If CData.Species(i).ThermoCoefficient.Count > 0 Then
                    t = CData.Species(i).CpValueFromShomateEquation(NumericUpDown1.Value)
                    o &= ToCSVItemFormat(t) & ","
                ElseIf CData.Species(i).CpData.Count > 0 Then
                    If CData.Species(i).CheckTemperatureValidityFromDataPoint(NumericUpDown1.Value) > -1 Then
                        t = CData.Species(i).CpValueFromAllDataPoint(NumericUpDown1.Value)
                        o &= ToCSVItemFormat(t) & ","
                    Else
                        o &= ","
                    End If
                Else
                    o &= ","
                End If
            End If
            If CheckedListBox1.GetItemChecked(5) Then
                Dim t As Double = CData.Species(i).EnthalpyValueFromShomateEquation(NumericUpDown1.Value)
                o &= ToCSVItemFormat(t) & ","
            End If
            If CheckedListBox1.GetItemChecked(6) Then
                Dim t As Double = CData.Species(i).EntropyValueFromShomateEquation(NumericUpDown1.Value)
                o &= ToCSVItemFormat(t) & ","
            End If
            If CheckedListBox1.GetItemChecked(7) Then
                Dim tempstr As String = ""
                If CData.Species(i).ThermoCoefficient.Count > 0 Then
                    For j As Integer = 0 To CData.Species(i).ThermoCoefficient.Count - 1
                        tempstr &= CData.Species(i).ThermoCoefficient(j).MinTemp & "-" & CData.Species(i).ThermoCoefficient(j).MaxTemp & "; "
                    Next
                    o &= tempstr & ","
                ElseIf CData.Species(i).CpData.Count > 0 Then
                    For j As Integer = 0 To CData.Species(i).CpData.Count - 1
                        tempstr &= "{"
                        CData.Species(i).CpData(j).Data.Sort(New Species.CpList.DataPointTemperatureComparer)
                        For k As Integer = 0 To CData.Species(i).CpData(j).Data.Count - 1
                            tempstr &= CData.Species(i).CpData(j).Data(k).Temperature & ";"
                        Next
                        tempstr &= "};"
                        'tempstr &= CData.Species(i).CpData(j).Data(0).Temperature & "-" & CData.Species(i).CpData(j).Data(CData.Species(i).CpData(j).Data.Count - 1).Temperature & "; "
                    Next
                    o &= tempstr & ","
                Else
                    o &= ","
                End If
            End If
            If CheckedListBox1.GetItemChecked(8) Then
                If CData.Species(i).ThermoCoefficient.Count > 0 Then
                    o &= "Shomate equation coefficient,"
                ElseIf CData.Species(i).CpData.Count > 0 Then
                    o &= "Data point,"
                Else
                    o &= ","
                End If
            End If
            If CheckedListBox1.GetItemChecked(9) Then
                Dim tmpstr As String = ""
                tmpstr &= CData.Species(i).EnthalpyStd298
                If CData.Species(i).EnthalpyStd298PlusMinus <> 0 Then tmpstr &= "±" & CData.Species(i).EnthalpyStd298PlusMinus
                o &= tmpstr & ","
                'If CData.Species(i).ThermoCoefficient.Count > 0 Then
                'Else
                '    o &= ","
                'End If
            End If
            If CheckedListBox1.GetItemChecked(10) Then
                Dim tmpstr As String = ""
                tmpstr &= CData.Species(i).EntropyStd298
                If CData.Species(i).EntropyStd298PlusMinus <> 0 Then tmpstr &= "±" & CData.Species(i).EntropyStd298PlusMinus
                o &= tmpstr & ","
                'If CData.Species(i).ThermoCoefficient.Count > 0 Then
                'Else
                '    o &= ","
                'End If
            End If
            If CheckedListBox1.GetItemChecked(11) Then
                Dim tmpstr As String = ""
                If CData.Species(i).ThermoDataFileExist Then
                    tmpstr = "True"
                Else
                    tmpstr = "False"
                End If
                o &= tmpstr & ","
                'If CData.Species(i).ThermoCoefficient.Count > 0 Then
                'Else
                '    o &= ","
                'End If
            End If
            If CheckedListBox1.GetItemChecked(12) Then
                Dim tmpstr As String = ""
                For tempt As Double = 50 To 2950 Step 50
                    tmpstr &= CData.Species(i).CpValueFromFitting(tempt) & ","
                Next
                tmpstr &= CData.Species(i).CpValueFromFitting(3000)

                o &= tmpstr & ","
                'If CData.Species(i).ThermoCoefficient.Count > 0 Then
                'Else
                '    o &= ","
                'End If
            End If
            o = Mid(o, 1, o.Length - 1) & vbCrLf
            If o.Length > 10000 Then
                My.Computer.FileSystem.WriteAllText(TextBox1.Text, o, True)
                o = ""
            End If
        Next
        My.Computer.FileSystem.WriteAllText(TextBox1.Text, o, True)
        MessageBox.Show("Finished!")
    End Sub
    Public Function ToCSVItemFormat(s As String) As String
        If s.Contains("""") Then
            s = s.Replace("""", """""")
            s = """" & s & """"
        ElseIf s.Contains(",") Then
            s = """" & s & """"
        End If
        Return s
    End Function
End Class