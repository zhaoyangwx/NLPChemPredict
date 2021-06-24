Imports System.IO
Imports hrefDownloader
Imports hrefDownloader.NumLib
Public Class Form1
    Public hList As List(Of PageItem)
    Public sList As List(Of PageItem)
    Public SpeciesList() As List(Of Integer）
    Public SpeciesInfo() As Species
    Public Class PIComparer
        Implements IEqualityComparer(Of PageItem)
        Public Overloads Function Equals(a As PageItem, b As PageItem) As Boolean Implements IEqualityComparer(Of PageItem).Equals
            If a Is Nothing And b Is Nothing Then Return True
            If a Is Nothing Or b Is Nothing Then Return False
            Return a.href = b.href
        End Function
        Public Overloads Function GetHashCode(o As PageItem) As Integer Implements IEqualityComparer(Of PageItem).GetHashCode
            Return o.href.GetHashCode
        End Function
    End Class


    Public Sub LoadList()
        Dim s As String = TextBox1.Text
        hList = New List(Of PageItem)
        Dim i As Integer = 0
        While i < s.Length - 8
            If Mid(s, i + 1, 8) = "a href=""" Then
                i += 8
                Dim j As Integer = i + 1
                While s(j) <> """" And j < s.Length
                    j += 1
                End While
                j -= 1
                Dim k As Integer = j + 1
                While Mid(s, k + 1, 9) <> "</td><td>" And k < s.Length - 9
                    k += 1
                End While
                k += 9
                Dim l As Integer = k + 1
                While Mid(s, l + 1, 5) <> "</td>" And l < s.Length - 5
                    l += 1
                End While
                l -= 1
                hList.Add(New PageItem(TextBox3.Text & Mid(s, i + 1, j - i + 1), Mid(s, k + 1, l - k + 1)))
                i = l
            End If
            i += 1
        End While
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LoadList()
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "href.txt", "", False)
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "Title.txt", "", False)

        Dim o As String = "", p As String = ""
        For i As Integer = 0 To hList.Count - 1
            o &= hList(i).Title & vbCrLf
            p &= hList(i).href & vbCrLf
            If o.Length > 10000 Then
                My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "Title.txt", o, True)
                o = ""
            End If
            If p.Length > 10000 Then
                My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "href.txt", p, True)
                p = ""
            End If
        Next
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "Title.txt", o, True)
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "href.txt", p, True)
    End Sub

    Private Sub Button2_Click(snd As Object, e As EventArgs) Handles Button2.Click
        'Net.ServicePointManager.ServerCertificateValidationCallback = New Net.Security.RemoteCertificateValidationCallback(
        '    Function(sender As Object,
        '             certificate As System.Security.Cryptography.X509Certificates.X509Certificate,
        '             chain As System.Security.Cryptography.X509Certificates.X509Chain,
        '             sslPolicyErrors As System.Net.Security.SslPolicyErrors)
        '        Return True
        '    End Function)
        Static Enabled As Boolean = True
        If Not Enabled Then Exit Sub
        Enabled = False
        LoadList()
        Dim progval As Integer = 0
        Dim digit As Integer = Digitof(hList.Count - 1)

        Dim th As New Threading.Thread(
            Sub()
                Parallel.For(0, hList.Count, New ParallelOptions With {.MaxDegreeOfParallelism = 6},
                    Sub(i As Integer)
                        If My.Computer.FileSystem.FileExists(TextBox2.Text & "\" & NumFormat(i, digit) & ".html") Then
                            If My.Computer.FileSystem.GetFileInfo(TextBox2.Text & "\" & NumFormat(i, digit) & ".html").Length > 0 Then
                                progval += 1
                                Exit Sub
                            End If
                        End If
                        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & NumFormat(i, digit) & "-Info.txt", hList(i).Title & vbCrLf & hList(i).href, False)
                        'Dim wc As New Net.WebClient
                        'Dim mycache As New Net.CredentialCache
                        ''wc.Credentials = New Net.NetworkCredential("", "")

                        'wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36")
                        ''wc.Credentials = New Net.NetworkCredential("", "", "kinetics.nist.gov")
                        'wc.DownloadFile(hList(i).href, TextBox2.Text & "\" & NumFormat(i, digit) & ".html")
                        Dim Completed As Boolean = False
                        Invoke(Sub()
                                   Dim wb As New WebBrowser With {.Parent = Me, .Left = 12, .Top = 238, .Width = 776, .Height = 186}
                                   AddHandler wb.DocumentCompleted,
                                       Sub()
                                           My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & NumFormat(i, digit) & ".html", wb.DocumentText, False)
                                           Completed = True
                                           wb.Dispose()
                                       End Sub
                                   wb.Navigate(hList(i).href)
                               End Sub)
                        While Not Completed
                            Threading.Thread.Sleep(100)
                        End While
                        progval += 1
                    End Sub)
                Enabled = True
                Invoke(Sub() MessageBox.Show("Finished!"))
            End Sub)
        Dim thprog As New Threading.Thread(
            Sub()
                While Not Enabled
                    Invoke(Sub() Text = progval & " / " & hList.Count)
                    Threading.Thread.Sleep(100)
                End While
            End Sub)
        thprog.Start()
        th.Start()
    End Sub
    Dim Prog3Th As New Action(
        Sub()
            Dim f() As IO.FileInfo = My.Computer.FileSystem.GetDirectoryInfo(TextBox2.Text & "\Reaction").GetFiles
            Dim fIDMaximum As Integer = f.Count \ 2 - 1
            Dim digit As Integer = Digitof(fIDMaximum)
            For i As Integer = 0 To fIDMaximum
                If Not My.Computer.FileSystem.FileExists(TextBox2.Text & "\Reaction" & "\" & NumFormat(i, digit) & ".html") Then
                    fIDMaximum = i - 1
                    Exit For
                End If
            Next
            If fIDMaximum < 0 Then Exit Sub

            ReDim SpeciesList(fIDMaximum)
            sList = New List(Of PageItem)
            totalprog = fIDMaximum + 1
            For i As Integer = 0 To fIDMaximum
                SpeciesList(i) = New List(Of Integer)
                Dim s As String = My.Computer.FileSystem.ReadAllText(TextBox2.Text & "\Reaction" & "\" & NumFormat(i, digit) & ".html")
                Dim t As String = "<div align=""center""><font size=""+3"">"
                Dim a As Integer = 0, b As Integer = 0
                For k As Integer = 0 To s.Length - 1 - t.Length
                    If Mid(s, k + 1, t.Length) = t Then
                        a = k + t.Length
                        Exit For
                    End If
                Next
                t = "</div>"
                For k As Integer = a + 1 To s.Length - 1 - t.Length
                    If Mid(s, k + 1, t.Length) = t Then
                        b = k - 1
                        Exit For
                    End If
                Next
                t = "<a href="""
                For k As Integer = a To b
                    If Mid(s, k + 1, t.Length) = t Then
                        Dim aa As Integer = k + t.Length
                        Dim bb As Integer
                        For bb = aa + 1 To b
                            If s(bb) = """" Then
                                bb -= 1
                                Exit For
                            End If
                        Next
                        Dim cc, dd As Integer
                        For cc = bb + 1 To b
                            If s(cc) = ">" Then
                                cc += 1
                                Exit For
                            End If
                        Next
                        Dim t2 As String = "</a>"
                        For kk As Integer = cc + 1 To b - t2.Length
                            If Mid(s, kk + 1, t2.Length) = t2 Then
                                dd = kk - 1
                                Exit For
                            End If
                        Next
                        Dim pi As New PageItem(Mid(s, aa + 1, bb - aa + 1), Mid(s, cc + 1, dd - cc + 1))
                        If Not sList.Contains(pi, New PIComparer) Then
                            sList.Add(pi)
                            SpeciesList(i).Add(sList.Count - 1)
                        Else
                            Dim id As Integer = -1
                            For id = 0 To sList.Count - 1
                                If PageItem.Equals(sList(id), pi) Then
                                    SpeciesList(i).Add(id)
                                    Exit For
                                End If
                            Next

                        End If
                    End If
                Next
                Threading.Interlocked.Add(progval, 1)
            Next
            Dim o As String = ""
            My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "Species.txt", o, False)
            totalprog = sList.Count
            progval = 0
            For i As Integer = 0 To sList.Count - 1
                o &= sList(i).Title & vbCrLf & sList(i).href & vbCrLf
                If o.Length > 10000 Then
                    My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "Species.txt", o, True)
                    o = ""
                End If
                Threading.Interlocked.Add(progval, 1)
            Next
            My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\" & "Species.txt", o, True)
            o = ""
            My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ReactionSpeciesID.txt", o, False)
            For i As Integer = 0 To fIDMaximum
                o &= i & vbTab
                For Each n As Integer In SpeciesList(i)
                    o &= n & vbTab
                Next
                o &= vbCrLf
                If o.Length > 10000 Then
                    My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ReactionSpeciesID.txt", o, True)
                    o = ""
                End If
            Next
            My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ReactionSpeciesID.txt", o, True)
        End Sub)
    Dim progval As Integer = 0
    Dim totalprog As Integer = 0

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Static IsBusy As Boolean = False
        If IsBusy Then Exit Sub
        IsBusy = True
        progval = 0
        totalprog = 0
        Dim th As New Threading.Thread(
            Sub()
                Prog3Th()
                IsBusy = False
                Invoke(Sub() MessageBox.Show("OK!"))
            End Sub)
        Dim thprog As New Threading.Thread(
            Sub()
                While IsBusy
                    Invoke(Sub() Text = progval & "/" & totalprog)
                    Threading.Thread.Sleep(100)
                End While
                Invoke(Sub() Text = progval & "/" & totalprog)
            End Sub)
        thprog.Start()
        th.Start()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Static Enabled As Boolean = True
        If Not Enabled Then Exit Sub
        Enabled = False
        LoadList()
        Dim progval As Integer = 0
        Dim digit As Integer = Digitof(sList.Count - 1)

        Dim th As New Threading.Thread(
            Sub()
                Parallel.For(0, sList.Count, New ParallelOptions With {.MaxDegreeOfParallelism = 6},
                    Sub(i As Integer)
                        If My.Computer.FileSystem.FileExists(TextBox2.Text & "\Species" & "\" & NumFormat(i, digit) & ".html") Then
                            If My.Computer.FileSystem.GetFileInfo(TextBox2.Text & "\Species" & "\" & NumFormat(i, digit) & ".html").Length > 0 Then
                                progval += 1
                                Exit Sub
                            End If
                        End If
                        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\Species" & "\" & NumFormat(i, digit) & "-Info.txt", sList(i).Title & vbCrLf & sList(i).href, False)
                        'Dim wc As New Net.WebClient
                        'Dim mycache As New Net.CredentialCache
                        ''wc.Credentials = New Net.NetworkCredential("", "")

                        'wc.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36")
                        ''wc.Credentials = New Net.NetworkCredential("", "", "kinetics.nist.gov")
                        'wc.DownloadFile(sList(i).href, TextBox2.Text & "\" & NumFormat(i, digit) & ".html")
                        Dim Completed As Boolean = False
                        Invoke(Sub()
                                   Dim wb As New WebBrowser With {.Parent = Me, .Left = 12, .Top = 238, .Width = 776, .Height = 186}
                                   AddHandler wb.DocumentCompleted,
                                       Sub()
                                           My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\Species" & "\" & NumFormat(i, digit) & ".html", wb.DocumentText, False)
                                           Completed = True
                                           wb.Dispose()
                                       End Sub
                                   wb.Navigate(sList(i).href)
                               End Sub)
                        While Not Completed
                            Threading.Thread.Sleep(100)
                        End While
                        progval += 1
                    End Sub)
                Enabled = True
                Invoke(Sub() MessageBox.Show("Finished!"))
            End Sub)
        Dim thprog As New Threading.Thread(
            Sub()
                While Not Enabled
                    Invoke(Sub() Text = progval & " / " & sList.Count)
                    Threading.Thread.Sleep(100)
                End While
                Invoke(Sub() Text = progval & " / " & sList.Count)
            End Sub)
        thprog.Start()
        th.Start()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Static Enabled As Boolean = True
        If Not Enabled Then Exit Sub
        Enabled = False
        Dim srcpath As String = TextBox2.Text & "\Species"
        Dim savepath As String = TextBox2.Text & "\ThermoData"
        Dim f() As IO.FileInfo = My.Computer.FileSystem.GetDirectoryInfo(srcpath).GetFiles
        Dim pcount As Integer = f.Count \ 2
        Dim digit As Integer = Digitof(pcount - 1)
        Dim progval As Integer = 0
        Dim totalprog As Integer = 0
        Dim th As New Threading.Thread(
            Sub()
                Dim hreflist As New List(Of PageItem)
                totalprog = pcount
                Parallel.For(0, pcount, New ParallelOptions With {.MaxDegreeOfParallelism = 6},
                    Sub(i As Integer)

                        'For i As Integer = 0 To pcount - 1
                        If Not My.Computer.FileSystem.FileExists(srcpath & "\" & NumFormat(i, digit) & ".html") Then
                            Exit Sub
                        End If
                        Dim s As String = My.Computer.FileSystem.ReadAllText(srcpath & "\" & NumFormat(i, digit) & ".html")
                        Dim target2 As String = """>Gas phase thermochemistry data"
                        Dim target1 As String = "<a href="""
                        Dim a, b As Integer
                        b = FindNextTextPosition(s, target2, 0) - 1
                        If b = -2 Then Exit Sub
                        'For b = 0 To s.Length - 1 - target2.Length
                        '    If Mid(s, b + 1, target2.Length) = target2 Then
                        '        b -= 1
                        '        Exit For
                        '    End If
                        'Next
                        'If Mid(s, b + 2, target2.Length) <> target2 Then
                        '    Exit Sub
                        'End If
                        a = FindPrevTextPosition(s, target1, b - 1 - target1.Length, True)
                        If a = -1 Then Exit Sub
                        'For a = Math.Max(0, b - 1 - target1.Length) To 0 Step -1
                        '    If Mid(s, a + 1, target1.Length) = target1 Then
                        '        a += target1.Length
                        '        Exit For
                        '    End If
                        'Next
                        Dim href As String = Mid(s, a + 1, b - a + 1)
                        href = "https://webbook.nist.gov/" & href
                        hreflist.Add(New PageItem(href, ""))
                        If My.Computer.FileSystem.FileExists(TextBox2.Text & "\ThermoData" & "\" & NumFormat(i, digit) & ".html") Then
                            If My.Computer.FileSystem.GetFileInfo(TextBox2.Text & "\ThermoData" & "\" & NumFormat(i, digit) & ".html").Length > 3500 Then
                                progval += 1
                                Exit Sub
                            End If
                        End If
                        Dim Completed As Boolean = False
                        Invoke(Sub()
                                   Dim wb As New WebBrowser With {.Parent = Me, .Left = 12, .Top = 238, .Width = 776, .Height = 186}
                                   AddHandler wb.DocumentCompleted,
                                       Sub()
                                           My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ThermoData" & "\" & NumFormat(i, digit) & ".html", wb.DocumentText, False)
                                           Completed = True
                                           wb.Dispose()
                                       End Sub
                                   wb.Navigate(href)
                               End Sub)
                        While Not Completed
                            Threading.Thread.Sleep(100)
                        End While
                        Threading.Interlocked.Add(progval, 1)
                        'Next
                    End Sub)
                'progval = 0
                'totalprog = hreflist.Count
                'Parallel.For(0, hreflist.Count,
                '    Sub(i As Integer)
                '        If My.Computer.FileSystem.FileExists(TextBox2.Text & "\ThermoData" & "\" & NumFormat(i, digit) & ".html") Then
                '            If My.Computer.FileSystem.GetFileInfo(TextBox2.Text & "\ThermoData" & "\" & NumFormat(i, digit) & ".html").Length > 3500 Then
                '                progval += 1
                '                Exit Sub
                '            End If
                '        End If
                '        Dim Completed As Boolean = False
                '        Invoke(Sub()
                '                   Dim wb As New WebBrowser With {.Parent = Me, .Left = 12, .Top = 238, .Width = 776, .Height = 186}
                '                   AddHandler wb.DocumentCompleted,
                '                                                  Sub()
                '                                                      My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ThermoData" & "\" & NumFormat(i, digit) & ".html", wb.DocumentText, False)
                '                                                      Completed = True
                '                                                      wb.Dispose()
                '                                                  End Sub
                '                   wb.Navigate(hreflist(i).href)
                '               End Sub)
                '        While Not Completed
                '            Threading.Thread.Sleep(100)
                '        End While
                '        Threading.Interlocked.Add(progval, 1)
                '    End Sub)
                MessageBox.Show("Finished!")
                Enabled = True
            End Sub)
        Dim thprog As New Threading.Thread(
            Sub()
                While Not Enabled
                    Invoke(Sub() Text = progval & " / " & totalprog)
                    Threading.Thread.Sleep(100)
                End While
                Invoke(Sub() Text = progval & " / " & totalprog)
            End Sub)
        thprog.Start()
        th.Start()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim Dir As String = TextBox2.Text & "\Reaction"
        Dim f() As IO.FileInfo = My.Computer.FileSystem.GetDirectoryInfo(Dir).GetFiles
        Dim fcount As Integer = f.Count \ 2
        Dim fdigit As Integer = Digitof(fcount - 1)
        Dim o As String = ""
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ReactionList.html", o, False)

        For i As Integer = 0 To fcount - 1
            Dim s As String = My.Computer.FileSystem.ReadAllText(Dir & "\" & NumFormat(i, fdigit) & "-Info.txt")
            Dim ss() As String = s.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            o &= "<a href=""" & ss(1) & """>" & ss(0) & "</a><br>"
            If o.Length > 10000 Then
                My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ReactionList.html", o, True)
                o = ""
            End If
        Next
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\ReactionList.html", o, True)
        MessageBox.Show("Finished!")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim ReactionCount As Integer = My.Computer.FileSystem.GetDirectoryInfo(TextBox2.Text & "\Reaction").GetFiles.Count \ 2
        Dim ReactionCountDigit As Integer = Digitof(ReactionCount - 1)
        Dim SpeciesCount As Integer = My.Computer.FileSystem.GetDirectoryInfo(TextBox2.Text & "\Species").GetFiles.Count \ 2
        Dim SpeciesCountDigit As Integer = Digitof(SpeciesCount - 1)
        ReDim SpeciesInfo(SpeciesCount - 1)
        For i As Integer = 0 To SpeciesCount - 1
            Dim s() As String = My.Computer.FileSystem.ReadAllText(TextBox2.Text & "\Species\" & NumFormat(i, SpeciesCountDigit) & "-Info.txt").Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            SpeciesInfo(i) = New Species
            SpeciesInfo(i).ReactionFormula = s(0)
            Dim htmlstr As String = My.Computer.FileSystem.ReadAllText(TextBox2.Text & "\Species\" & NumFormat(i, SpeciesCountDigit) & ".html")
            Dim p As Integer = FindNextTextPosition(htmlstr, "<title>", 0, True)
            If p = -1 Then Continue For
            Dim q As Integer = FindNextTextPosition(htmlstr, "</title>", p) - 1
            If p <> -1 Then
                SpeciesInfo(i).Name = GetSubString(htmlstr, p, q)
            End If
            p = FindNextTextPosition(htmlstr, ">Formula</a>:</strong> ", q + 1, True)
            If p = -1 Then p = FindNextTextPosition(htmlstr, ">Formula</a>:</strong>", q + 1, True)
            q = FindNextTextPosition(htmlstr, "</li>", p + 1) - 1
            If p <> -1 Then
                SpeciesInfo(i).Formula = GetSubString(htmlstr, p, q)
            End If
            If My.Computer.FileSystem.FileExists(TextBox2.Text & "\ThermoData\" & NumFormat(i, SpeciesCountDigit) & ".html") Then
                SpeciesInfo(i).ThermoDataFileExist = True
                htmlstr = My.Computer.FileSystem.ReadAllText(TextBox2.Text & "\ThermoData\" & NumFormat(i, SpeciesCountDigit) & ".html")
                p = FindNextTextPosition(htmlstr, "Gas Phase Heat Capacity (Shomate Equation)", 0, True)
                If p <> -1 Then
                    p = FindNextTextPosition(htmlstr, "Temperature (K)", p, True)
                    Dim tmp1 As Integer = FindNextTextPosition(htmlstr, "<td", p)
                    Dim tmp2 As Integer = FindNextTextPosition(htmlstr, "</tr>", p, True)

                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    Dim t1 As String = GetSubString(htmlstr, p, q)
                    Dim val2 As Integer = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q, True)
                    Dim tdcount As Integer = (GetSubString(htmlstr, p, tmp2).Length - GetSubString(htmlstr, p, tmp2).Replace("</td>", "").Length) / 5
                    If val2 > tmp2 Then
                        tdcount = 1
                    Else
                        'p = val2
                    End If
                    'MessageBox.Show(tdcount & " " & i)
                    Dim TCoeff(tdcount - 1) As Species.ShomateCoeff
                    TCoeff(0) = New Species.ShomateCoeff
                    TCoeff(0).SetTemperatureRange(t1)
                    q = p - 1

                    For ii As Integer = 1 To tdcount - 1
                        p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q, True)
                        q = FindNextTextPosition(htmlstr, "</td>", p) - 1
                        Dim t2 As String = GetSubString(htmlstr, p, q)
                        TCoeff(ii) = New Species.ShomateCoeff
                        TCoeff(ii).SetTemperatureRange(t2)
                    Next

                    p = FindNextTextPosition(htmlstr, ">A</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).A = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).A = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">B</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).B = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).B = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">C</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).C = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).C = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">D</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).D = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).D = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">E</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).E = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).E = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">F</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).F = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).F = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">G</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).G = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).G = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    p = FindNextTextPosition(htmlstr, ">H</th>", p + 1)
                    p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", p + 1, True)
                    q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                    TCoeff(0).H = TryVal(GetSubString(htmlstr, p, q))
                    If tdcount > 0 Then
                        For ii As Integer = 1 To tdcount - 1
                            p = FindNextTextPosition(htmlstr, "<td class=""exp right-nowrap"">", q + 1, True)
                            q = FindNextTextPosition(htmlstr, "</td>", p + 1) - 1
                            TCoeff(ii).H = TryVal(GetSubString(htmlstr, p, q))
                        Next
                    End If

                    For ii As Integer = 0 To tdcount - 1
                        SpeciesInfo(i).ThermoCoefficient.Add(TCoeff(ii))
                    Next
                End If

                p = FindNextTextPosition(htmlstr, "Constant pressure heat capacity of gas", 0, True)
                While p <> -1
                    Dim pmax As Integer = FindNextTextPosition(htmlstr, "</table>", p)
                    Dim tlst As New Species.CpList
                    While FindNextTextPosition(htmlstr, "<td class=""right-nowrap"">", p, True) < pmax And FindNextTextPosition(htmlstr, "<td class=""right-nowrap"">", p, True) <> -1
                        Dim a As Integer = FindNextTextPosition(htmlstr, "<td class=""right-nowrap"">", p, True)
                        Dim b As Integer = FindNextTextPosition(htmlstr, "</td>", a) - 1
                        Dim c As Integer = FindNextTextPosition(htmlstr, "<td class=""right-nowrap"">", b, True)
                        Dim d As Integer = FindNextTextPosition(htmlstr, "</td>", c) - 1
                        p = d + 1
                        Dim val1() As String = GetSubString(htmlstr, a, b).Split({"&plusmn;"}, StringSplitOptions.RemoveEmptyEntries)
                        Dim val2() As String = GetSubString(htmlstr, c, d).Split({"&plusmn;"}, StringSplitOptions.RemoveEmptyEntries)
                        If val1.Length = 1 Then
                            ReDim Preserve val1(1)
                            val1(1) = "0"
                        End If
                        If val2.Length = 1 Then
                            ReDim Preserve val2(1)
                            val2(1) = "0"
                        End If
                        tlst.Add(TryVal(val2(0)), TryVal(val1(0)), TryVal(val2(1)), TryVal(val1(1)))
                    End While
                    If tlst.Data.Count > 0 Then SpeciesInfo(i).CpData.Add(tlst)
                    p = FindNextTextPosition(htmlstr, "Constant pressure heat capacity of gas", p, True)
                End While
                p = FindNextTextPosition(htmlstr, "<sub>f</sub>H&deg;<sub>gas</sub></td><td class=""right-nowrap"">", 0, True)
                If p <> -1 Then
                    q = FindNextTextPosition(htmlstr, "</td>", p, False) - 1
                    Dim val0() As String = GetSubString(htmlstr, p, q).Split({"&plusmn;"}, StringSplitOptions.RemoveEmptyEntries)
                    If val0.Length = 1 Then
                        ReDim Preserve val0(1)
                        val0(1) = "0"
                    End If
                    SpeciesInfo(i).EnthalpyStd298 = TryVal(val0(0))
                    SpeciesInfo(i).EnthalpyStd298PlusMinus = TryVal(val0(1))
                End If
                p = FindNextTextPosition(htmlstr, "<td style=""text-align: left;"">S&deg;<sub>gas,1 bar</sub></td><td class=""right-nowrap"">", 0, True)
                If p <> -1 Then
                    q = FindNextTextPosition(htmlstr, "</td>", p, False) - 1
                    Dim val0() As String = GetSubString(htmlstr, p, q).Split({"&plusmn;"}, StringSplitOptions.RemoveEmptyEntries)
                    If val0.Length = 1 Then
                        ReDim Preserve val0(1)
                        val0(1) = "0"
                    End If
                    SpeciesInfo(i).EntropyStd298 = TryVal(val0(0))
                    SpeciesInfo(i).EntropyStd298PlusMinus = TryVal(val0(1))
                    'If SpeciesInfo(i).ThermoCoefficient.Count > 0 Then
                    'End If
                Else
                    p = FindNextTextPosition(htmlstr, "<td style=""text-align: left;"">S&deg;<sub>gas</sub></td><td class=""right-nowrap"">", 0, True)
                    If p <> -1 Then
                        q = FindNextTextPosition(htmlstr, "</td>", p, False) - 1
                        Dim val0() As String = GetSubString(htmlstr, p, q).Split({"&plusmn;"}, StringSplitOptions.RemoveEmptyEntries)
                        If val0.Length = 1 Then
                            ReDim Preserve val0(1)
                            val0(1) = "0"
                        End If
                        SpeciesInfo(i).EntropyStd298 = TryVal(val0(0))
                        SpeciesInfo(i).EntropyStd298PlusMinus = TryVal(val0(1))
                        'If SpeciesInfo(i).ThermoCoefficient.Count > 0 Then
                        'End If
                    End If
                End If
            End If
        Next
        MessageBox.Show("Finished!")
    End Sub
    Public Function TryVal(s As String) As Double
        If s Is Nothing Then s = ""
        If s = "" Then Return 0
        s = s.Replace("&times;10<sup>", "e")
        s = s.Replace("</sup>", "")
        s = s.Replace(" ", "")
        Return Val(s)
    End Function
    Public Function FindNextTextPosition(ByRef Source As String, ByRef Target As String, ByVal StartPosition As Integer, Optional ByVal ExcludedPosition As Boolean = False) As Integer
        If Target = "" Then Return -1
        For i As Integer = StartPosition To Source.Length - Target.Length
            If Mid(Source, i + 1, Target.Length) = Target Then
                If Not ExcludedPosition Then
                    Return i
                Else
                    Return i + Target.Length
                End If
            End If
        Next
        Return -1
    End Function
    Public Function FindPrevTextPosition(ByRef Source As String, ByRef Target As String, ByVal StartPosition As Integer, Optional ByVal ExcludedPosition As Boolean = False) As Integer
        If Target = "" Then Return -1
        For i As Integer = Math.Min(StartPosition, Source.Length - Target.Length) To 0 Step -1
            If Mid(Source, i + 1, Target.Length) = Target Then
                If Not ExcludedPosition Then
                    Return i
                Else
                    Return i + Target.Length
                End If
            End If
        Next
        Return -1
    End Function

    Public Function GetSubString(ByRef Source As String, ByVal a As Integer, ByVal b As Integer) As String
        If a > b Then
            Dim t As Integer = a
            a = b
            b = t
        End If
        If a < 0 Then a = 0
        If b > Source.Length - 1 Then b = Source.Length - 1
        Return Mid(Source, a + 1, b - a + 1)
    End Function
    Public Output As ChemData
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Output = New ChemData
        Output.Species = SpeciesInfo.ToList
        Output.ReactionSpeciesID = SpeciesList.ToList
        Output.LoadReactionList(TextBox2.Text & "\Reaction")
        My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\Result.xml", Output.GetSerializedText, False)
        MessageBox.Show("Finished!")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Static IsBusy As Boolean = False
        If IsBusy Then Exit Sub
        IsBusy = True
        progval = 0
        totalprog = 0
        Dim th As New Threading.Thread(
            Sub()
                If SpeciesList Is Nothing Then
                    Prog3Th()
                End If
                If SpeciesInfo Is Nothing Then
                    Button7_Click(sender, e)
                End If
                If Output Is Nothing Then
                    Output = New ChemData
                    Output.Species = SpeciesInfo.ToList
                    Output.ReactionSpeciesID = SpeciesList.ToList
                    Output.LoadReactionList(TextBox2.Text & "\Reaction")
                End If
                Dim frm2 As New Form2 With {.CData = Output}
                Invoke(Sub() frm2.Show()）
                IsBusy = False
                Invoke(Sub() MessageBox.Show("Ready!"))
            End Sub)
        Dim thprog As New Threading.Thread(
            Sub()
                While IsBusy
                    Invoke(Sub() Text = progval & "/" & totalprog)
                    Threading.Thread.Sleep(100)
                End While
                Invoke(Sub() Text = progval & "/" & totalprog)
            End Sub)
        thprog.Start()
        th.Start()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click
        Exit Sub
        Dim s As String = My.Computer.FileSystem.ReadAllText("D:\Documents\Visual Studio 2017\Projects\FasttextTest\FasttextTest\train_unsupervised_alph.txt")
        s = s.Replace(" ", "_")
        Dim sb As New System.Text.StringBuilder
        For i As Integer = 0 To s.Length - 1
            sb.Append(s(i))
            sb.Append(" ")
        Next
        s = sb.ToString
        My.Computer.FileSystem.WriteAllText("D:\Documents\Visual Studio 2017\Projects\FasttextTest\FasttextTest\train_unsupervised_sep.txt", s, False)
    End Sub
    Public Function NumFormatConv36(decval As Integer, digit As Integer) As String
        Dim out As String = ""
        While decval > 0
            Dim d As Integer = decval Mod 36
            If d < 10 Then
                out = d.ToString & out
            Else
                out = Chr(d - 10 + Asc("a")) & out
            End If
            decval \= 36
        End While
        While out.Length < digit
            out = "0" & out
        End While
        Return out
    End Function
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Static Enabled As Boolean = True
        If Not Enabled Then Exit Sub
        Enabled = False
        LoadList()
        Dim progval As Integer = 0
        Dim pagecount As Integer = 1240096
        Dim digit As Integer = 4
        Dim NewFileCount As Integer = 0
        Dim th As New Threading.Thread(
            Sub()
                Parallel.For(0, pagecount, New ParallelOptions With {.MaxDegreeOfParallelism = 6},
                    Sub(i As Integer)
                        If My.Computer.FileSystem.FileExists(TextBox2.Text & "\MOLInstinctsData\" & NumFormatConv36(i, digit) & ".html") Then
                            If My.Computer.FileSystem.GetFileInfo(TextBox2.Text & "\MOLInstinctsData\" & NumFormatConv36(i, digit) & ".html").Length > 900 Then
                                Threading.Interlocked.Add(progval, 1)
                                Exit Sub
                            End If
                        End If
                        Dim Completed As Boolean = False
                        Invoke(Sub()
                                   Dim wb As New WebBrowser With {.Parent = Me, .Left = 12, .Top = 238, .Width = 776, .Height = 186, .ScriptErrorsSuppressed = True, .Visible = False}
                                   AddHandler wb.DocumentCompleted,
                                       Sub()
                                           My.Computer.FileSystem.WriteAllText(TextBox2.Text & "\MOLInstinctsData\" & NumFormatConv36(i, digit) & ".html", wb.DocumentText, False)
                                           Completed = True
                                           Threading.Interlocked.Add(NewFileCount, 1)
                                           wb.Dispose()
                                       End Sub
                                   AddHandler wb.Navigated,
                                       Sub()
                                           Dim vDocument As mshtml.IHTMLDocument2 = CType(wb.Document.DomDocument, mshtml.IHTMLDocument2)
                                           vDocument.parentWindow.execScript("function confirm(str){return true;} ", "javascript")
                                           vDocument.parentWindow.execScript("function alert(str){return true;} ", "javaScript")
                                       End Sub
                                   wb.Navigate("http://search.imolinstincts.com.cn/properties/constantProperty.ce?0001-" & NumFormatConv36(i, digit))
                               End Sub)
                        While Not Completed
                            Threading.Thread.Sleep(100)
                        End While
                        Threading.Interlocked.Add(progval, 1)
                    End Sub)
                Enabled = True
                Invoke(Sub() MessageBox.Show("Finished!"))
            End Sub)
        Dim thprog As New Threading.Thread(
            Sub()
                While Not Enabled
                    Invoke(Sub() Text = progval & "(+" & NewFileCount & ") / " & pagecount)
                    Threading.Thread.Sleep(100)
                End While
            End Sub)
        thprog.Start()
        th.Start()
    End Sub
End Class
Public Class NumLib
    Public Shared Function NumFormat(value As Integer, Optional ByVal digit As Integer = 0) As String
        Dim s As String = value.ToString
        While s.Length < digit
            s = "0" & s
        End While
        Return s
    End Function
    Public Shared Function Digitof(ByVal n As Integer) As Integer
        Dim t As Integer = 0
        While n > 0
            n \= 10
            t += 1
        End While
        Return t
    End Function
    Public Shared Function Digitof(ByVal n As String) As Integer
        Return n.Length
    End Function

End Class
<Serializable>
Public Class PageItem
    Public Property href As String
    Public Property Title As String
    Public Sub New(hrefs As String, Optional ByVal Titles As String = "")
        href = hrefs.Replace("&amp;", "&")
        Title = Titles
    End Sub
    Public Sub New()
        href = ""
        Title = ""
    End Sub
    Public Shared Function Equals(a As PageItem, b As PageItem) As Boolean
        Return a.href = b.href
    End Function
End Class
<Serializable>
Public Class ChemData
    Public Property Species As List(Of Species)
    Public Property ReactionSpeciesID As List(Of List(Of Integer))
    Public Property Reaction As List(Of PageItem)
    Public Sub LoadReactionList(ByVal Dir As String)
        Dim fcount As Integer = My.Computer.FileSystem.GetDirectoryInfo(Dir).GetFiles.Count \ 2
        Dim fcdigit As Integer = Digitof(fcount - 1)
        Dim r(fcount - 1) As PageItem
        For i As Integer = 0 To fcount - 1
            Dim s() As String = My.Computer.FileSystem.ReadAllText(Dir & "\" & NumFormat(i, fcdigit) & "-Info.txt").Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            r(i) = New PageItem(s(1), s(0))
        Next
        Reaction = r.ToList
    End Sub
    Public Function GetSerializedText()
        Dim writer As New System.Xml.Serialization.XmlSerializer(GetType(ChemData))
        Dim sb As New System.Text.StringBuilder()
        Dim t As IO.TextWriter = New IO.StringWriter(sb)
        writer.Serialize(t, Me)
        t.Close()
        Return sb.ToString
    End Function
    Public Shared Function FromXML(s As String) As ChemData
        Dim reader As New System.Xml.Serialization.XmlSerializer(GetType(ChemData))
        Dim t As IO.TextReader = New IO.StringReader(s)
        Return CType(reader.Deserialize(t), ChemData)
    End Function
End Class
<Serializable>
Public Class Species
    Public Property Name As String
    Public Property ReactionFormula As String
    Public Property Formula As String
    Public Property EntropyStd298 As Double
    Public Property EntropyStd298PlusMinus As Double
    Private _EnthalpyStd298 As Double
    Private _EnthalpyStd298PlusMinus As Double
    Public Property ThermoDataFileExist As Boolean
    Public Property EnthalpyStd298 As Double
        Set(value As Double)
            If ThermoCoefficient.Count > 0 Then
                For Each tc As ShomateCoeff In ThermoCoefficient
                    tc.EnthalpyStd298 = value
                Next
            End If
            _EnthalpyStd298 = value
        End Set
        Get
            If ThermoCoefficient.Count > 0 Then
                Return ThermoCoefficient(0).EnthalpyStd298
            Else
                Return _EnthalpyStd298
            End If
        End Get
    End Property
    Public Property EnthalpyStd298PlusMinus As Double
        Set(value As Double)
            If ThermoCoefficient.Count > 0 Then
                For Each tc As ShomateCoeff In ThermoCoefficient
                    tc.EnthalpyStd298PlusMinus = value
                Next
            End If
            _EnthalpyStd298PlusMinus = value
        End Set
        Get
            If ThermoCoefficient.Count > 0 Then
                Return ThermoCoefficient(0).EnthalpyStd298PlusMinus
            Else
                Return _EnthalpyStd298PlusMinus
            End If
        End Get
    End Property
    <Serializable>
    Public Class ShomateCoeff
        Public Property MinTemp As Double
        Public Property MaxTemp As Double
        Public Property EnthalpyStd298 As Double
        Public Property EnthalpyStd298PlusMinus As Double
        Public Sub SetTemperatureRange(s As String)
            Dim t() As String = s.Split({"-"}, StringSplitOptions.RemoveEmptyEntries)
            If t.Length <> 2 Then Exit Sub
            t(0) = t(0).Replace(" ", "")
            t(1) = t(1).Replace(" ", "")
            MinTemp = Val(t(0))
            MaxTemp = Val(t(1))
        End Sub
        Public Function IsValidTemperature(Temperature As Double) As Boolean
            Return (Temperature >= MinTemp And Temperature <= MaxTemp)
        End Function
        Public Function GetCpValue(Temperature As Double) As Double 'Cp
            'If Temperature < MinTemp Or Temperature > MaxTemp Then Return 0
            If Temperature <= 0 Then Return 0
            Dim t As Double = Temperature / 1000
            Return A + B * t + C * t ^ 2 + D * t ^ 3 + E / (t ^ 2)
        End Function
        Public Function GetEnthalpyValue(Temperature As Double) As Double 'H
            If Temperature < MinTemp Or Temperature > MaxTemp Then Return 0
            If Temperature = 0 Then Return 0
            Dim t As Double = Temperature / 1000
            Return A * t + B * t ^ 2 / 2 + C * t ^ 3 / 3 + D * t ^ 4 / 4 - E / t + F - H + EnthalpyStd298
        End Function
        Public Function GetEntropyValue(Temperature As Double) As Double 'S
            If Temperature < MinTemp Or Temperature > MaxTemp Then Return 0
            If Temperature = 0 Then Return 0
            Dim t As Double = Temperature / 1000
            Return A * Math.Log(t) + B * t + C * t ^ 2 / 2 + D * t ^ 3 / 3 - E / (2 * t ^ 2) + G
        End Function


        Public Function GetCoeffList() As List(Of Double)
            Return _Coeff.ToList
        End Function

        Public Sub SetCoeff(Coeff As List(Of Double))
            For i As Integer = 0 To Math.Min(7, Coeff.Count - 1)
                _Coeff(i) = Coeff(i)
            Next
            If Coeff.Count < 8 Then
                For i As Integer = Coeff.Count To 7
                    _Coeff(i) = 0
                Next
            End If
        End Sub
        <NonSerialized>
        Private _Coeff() As Double
        Public Property A As Double
            Get
                Return _Coeff(0)
            End Get
            Set(value As Double)
                _Coeff(0) = value
            End Set
        End Property
        Public Property B As Double
            Get
                Return _Coeff(1)
            End Get
            Set(value As Double)
                _Coeff(1) = value
            End Set
        End Property
        Public Property C As Double
            Get
                Return _Coeff(2)
            End Get
            Set(value As Double)
                _Coeff(2) = value
            End Set
        End Property
        Public Property D As Double
            Get
                Return _Coeff(3)
            End Get
            Set(value As Double)
                _Coeff(3) = value
            End Set
        End Property
        Public Property E As Double
            Get
                Return _Coeff(4)
            End Get
            Set(value As Double)
                _Coeff(4) = value
            End Set
        End Property
        Public Property F As Double
            Get
                Return _Coeff(5)
            End Get
            Set(value As Double)
                _Coeff(5) = value
            End Set
        End Property
        Public Property G As Double
            Get
                Return _Coeff(6)
            End Get
            Set(value As Double)
                _Coeff(6) = value
            End Set
        End Property
        Public Property H As Double
            Get
                Return _Coeff(7)
            End Get
            Set(value As Double)
                _Coeff(7) = value
            End Set
        End Property
        Public Sub New()
            ReDim _Coeff(7)
        End Sub
        Public Sub New(ByVal Coeff As List(Of Double))
            ReDim _Coeff(7)
            SetCoeff(Coeff)
        End Sub
        Public Sub New(ByVal Coeff() As Double)
            ReDim _Coeff(7)
            For i As Integer = 0 To Math.Min(7, Coeff.Length - 1)
                _Coeff(i) = Coeff(i)
            Next
        End Sub
    End Class
    Public Property ThermoCoefficient As List(Of ShomateCoeff)
    <Serializable>
    Public Class CpList
        Public Structure DataPoint
            Public Temperature As Double
            Public TemperaturePlusMinus As Double
            Public Cp As Double
            Public CpPlusMinus As Double
            Public Sub New(ByVal temp As Double, ByVal value As Double, Optional ByVal temppm As Double = 0, Optional ByVal valuepm As Double = 0)
                Temperature = temp
                Cp = value
                TemperaturePlusMinus = temppm
                CpPlusMinus = valuepm
            End Sub
            Public Shared Widening Operator CType(o As DataPoint) As Double
                Return o.Cp
            End Operator
        End Structure
        Public Property Data As List(Of DataPoint)
        Public Sub Add(temp As Double, value As Double, Optional ByVal temppm As Double = 0, Optional ByVal valuepm As Double = 0)
            Data.Add(New DataPoint(temp, value, temppm, valuepm))
        End Sub
        Public Sub Add(DataPoints As CpList, Optional ByVal Overwrite As Boolean = False, Optional ByVal KeepBoth As Boolean = False)
            For i As Integer = 0 To DataPoints.Data.Count - 1
                If Not Overwrite Then
                    If Not KeepBoth Then
                        Dim containp As Boolean = False
                        For j As Integer = 0 To Data.Count - 1
                            If Data(j).Temperature = DataPoints.Data(i).Temperature Then
                                containp = True
                                Exit For
                            End If
                        Next
                        If containp Then Continue For
                        Data.Add(DataPoints.Data(i))
                    Else
                        Data.Add(DataPoints.Data(i))
                    End If
                Else
                    Dim containp As Boolean = False
                    For j As Integer = 0 To Data.Count - 1
                        If Data(j).Temperature = DataPoints.Data(i).Temperature Then
                            Data(j) = DataPoints.Data(i)
                            containp = True
                            Exit For
                        End If
                    Next
                    If Not containp Then Data.Add(DataPoints.Data(i))
                End If
            Next
        End Sub
        Public Sub New()
            Data = New List(Of DataPoint)
        End Sub
        Public Class DataPointEqualityComparer
            Implements IEqualityComparer(Of DataPoint)
            Public Overloads Function Equals(a As DataPoint, b As DataPoint) As Boolean Implements IEqualityComparer(Of DataPoint).Equals
                Return a.Temperature = b.Temperature And a.Cp = b.Cp And a.TemperaturePlusMinus = b.TemperaturePlusMinus And a.CpPlusMinus = b.CpPlusMinus
            End Function
            Public Overloads Function GetHashCode(o As DataPoint) As Integer Implements IEqualityComparer(Of DataPoint).GetHashCode
                Return {o.Temperature, o.TemperaturePlusMinus, o.Cp, o.CpPlusMinus}.GetHashCode
            End Function
        End Class
        Public DTempC As New DataPointTemperatureComparer
        Public Class DataPointTemperatureComparer
            Implements IComparer(Of DataPoint)
            Public Function Compare(x As DataPoint, y As DataPoint) As Integer Implements IComparer(Of DataPoint).Compare
                If x.Temperature < y.Temperature Then
                    Return -1
                ElseIf x.Temperature = y.Temperature Then
                    If x.TemperaturePlusMinus < y.TemperaturePlusMinus Then
                        Return -1
                    ElseIf x.TemperaturePlusMinus = y.TemperaturePlusMinus Then
                        If x.Cp < y.Cp Then
                            Return -1
                        ElseIf x.Cp = y.Cp Then
                            If x.CpPlusMinus < y.CpPlusMinus Then
                                Return -1
                            ElseIf x.CpPlusMinus = y.CpPlusMinus Then
                                Return 0
                            Else
                                Return 1
                            End If
                        Else
                            Return 1
                        End If
                    Else
                        Return 1
                    End If
                Else
                    Return 1
                End If
            End Function
        End Class
    End Class
    Public Property CpData As List(Of CpList)
    Public Class SPLine
        Public Xi, Yi, A, B, C, H, Lambda, Mu, G, M As Double()
        Public NN, n As Integer
        Public Sub New()
            NN = 0
            n = 0
        End Sub
        Public Function Init(ByVal Xi As Double(), ByVal Yi As Double()) As Boolean
            If Xi Is Nothing Or Yi Is Nothing Then Return False
            If Xi.Length <> Yi.Length Then Return False
            If Xi.Length = 0 Then Return False

            'Init NN size
            NN = Xi.Length
            n = NN - 1

            'Init coeff
            ReDim A(NN - 2)
            ReDim B(NN - 1)
            ReDim C(NN - 2)
            ReDim Me.Xi(NN - 1)
            ReDim Me.Yi(NN - 1)
            ReDim H(NN - 2)
            ReDim Lambda(NN - 2)
            ReDim Mu(NN - 2)
            ReDim G(NN - 1)
            ReDim M(NN - 1)

            'Load Points
            For i As Integer = 0 To n
                Me.Xi(i) = Xi(i)
                Me.Yi(i) = Yi(i)
            Next

            GetH()
            GetLambda_Mu_G()
            GetABC()

            Dim chase As New Chasing()
            chase.init(A, B, C, G)
            chase.Solve(M)
            Return True
        End Function
        Public Sub GetH()
            'Calc h
            For i As Integer = 0 To n - 1
                H(i) = Xi(i + 1) - Xi(i)
            Next
        End Sub
        Public Sub GetLambda_Mu_G()
            'Calc Lambda, Mu, G
            Dim t1, t2 As Double
            For i As Integer = 1 To n - 1
                Lambda(i) = H(i) / (H(i) + H(i - 1))
                Mu(i) = 1 - Lambda(i)
                t1 = (Yi(i) - Yi(i - 1)) / H(i - 1)
                t2 = (Yi(i + 1) - Yi(i)) / H(i)
                G(i) = 3 * (Lambda(i) * t1 + Mu(i) * t2)
            Next
            G(0) = 3 * (Yi(1) - Yi(0)) / H(0)
            G(n) = 3 * (Yi(n) - Yi(n - 1)) / H(n - 1)
            Mu(0) = 1
            Lambda(0) = 0
        End Sub
        Public Sub GetABC()
            'Calc A B C
            For i As Integer = 1 To n - 1
                A(i - 1) = Lambda(i)
                C(i) = Mu(i)
            Next
            A(n - 1) = 1
            C(0) = 1
            For i As Integer = 0 To n
                B(i) = 2
            Next
        End Sub
        Public Function fai0(x As Double) As Double
            Dim t1, t2, s As Double
            t1 = 2 * x + 1
            t2 = (x - 1) ^ 2
            s = t1 * t2
            Return s
        End Function
        Public Function fai1(x As Double) As Double
            Return x * (x - 1) ^ 2
        End Function
        Public Function Interpolate(x As Double) As Double
            Dim s As Double = 0
            Dim P1, P2 As Double
            Dim t As Double = x
            Dim iNum As Integer = GetSection(x)
            If iNum = -1 Then
                iNum = 0
                t = Xi(iNum)
                Return Yi(0)
            End If
            If iNum = -999 Then
                iNum = n - 1
                t = Xi(iNum + 1)
                Return Yi(n)
            End If
            P1 = (t - Xi(iNum)) / H(iNum)
            P2 = (Xi(iNum + 1) - t) / H(iNum)
            s = Yi(iNum) * fai0(P1) + Yi(iNum + 1) * fai0(P2) + M(iNum) * H(iNum) * fai1(P1) - M(iNum + 1) * H(iNum) * fai1(P2)
            Return s
        End Function
        Public Function GetSection(x As Double) As Integer
            Dim iNum As Integer = -1
            If x < Xi(0) Then Return -1
            If x > Xi(n - 1) Then Return -999
            For i As Integer = 0 To n - 1
                If x >= Xi(i) And x <= Xi(i + 1) Then
                    iNum = i
                    Exit For
                End If
            Next
            Return iNum
        End Function
        Public Class Chasing
            Public N As Integer
            Public d, Aa, Ab, Ac, L, U, S As Double()
            Public Function init(a As Double(), b As Double(), c As Double(), d As Double()) As Boolean
                Dim na As Integer = a.Length
                Dim nb As Integer = b.Length
                Dim nc As Integer = c.Length
                Dim nd As Integer = d.Length
                If nb < 3 Then Return False
                N = nb
                If na <> N - 1 Or nc <> N - 1 Or nd <> N Then Return False
                ReDim S(N - 1)
                ReDim L(N - 2)
                ReDim U(N - 1)

                ReDim Aa(N - 2)
                ReDim Ab(N - 1)
                ReDim Ac(N - 2)
                ReDim Me.d(N - 1)

                For i As Integer = 0 To N - 2
                    Ab(i) = b(i)
                    Me.d(i) = d(i)
                    Aa(i) = a(i)
                    Ac(i) = c(i)
                Next
                Ab(N - 1) = b(N - 1)
                Me.d(N - 1) = d(N - 1)
                Return True
            End Function
            Public Function Solve(ByRef R As Double()) As Boolean
                ReDim R(Ab.Length - 1)
                U(0) = Ab(0)
                For i As Integer = 2 To N
                    L(i - 2) = Aa(i - 2) / U(i - 2)
                    U(i - 1) = Ab(i - 1) - Ac(i - 2) * L(i - 2)
                Next
                Dim Y(d.Length - 1) As Double
                Y(0) = d(0)
                For i As Integer = 2 To N
                    Y(i - 1) = d(i - 1) - L(i - 2) * Y(i - 2)
                Next
                R(N - 1) = Y(N - 1) / U(N - 1)
                For i As Integer = N - 1 To 1 Step -1
                    R(i - 1) = (Y(i - 1) - Ac(i - 1) * R(i)) / U(i - 1)
                Next
                For i As Integer = 0 To R.Length - 1
                    If Double.IsInfinity(R(i)) Or Double.IsNaN(R(i)) Then Return False
                Next
                Return True
            End Function
        End Class
    End Class
    Public Function CpValueFromShomateEquation(ByVal Temperature As Double) As Double
        Dim tcomin As ShomateCoeff, deltatmin As Double = 99999

        For Each tco As ShomateCoeff In ThermoCoefficient
            If tco.IsValidTemperature(Temperature) Then
                Return tco.GetCpValue(Temperature)
            End If
            'Dim deltat As Double = Math.Min(Math.Abs(tco.MaxTemp - Temperature), Math.Abs(tco.MinTemp - Temperature))
            'If deltatmin > deltat Then
            '    deltatmin = deltat
            '    tcomin = tco
            'End If
        Next
        Return -1
        If tcomin IsNot Nothing Then Return tcomin.GetCpValue(Temperature)
        Return 0
    End Function
    Public Function EnthalpyValueFromShomateEquation(ByVal Temperature As Double) As Double
        For Each tco As ShomateCoeff In ThermoCoefficient
            If tco.IsValidTemperature(Temperature) Then
                Return tco.GetEnthalpyValue(Temperature)
            End If
        Next
        Return 0
    End Function
    Public Function EntropyValueFromShomateEquation(ByVal Temperature As Double) As Double
        For Each tco As ShomateCoeff In ThermoCoefficient
            If tco.IsValidTemperature(Temperature) Then
                Return tco.GetEntropyValue(Temperature)
            End If
        Next
        Return 0
    End Function
    Public Function CheckTemperatureValidityFromDataPoint(ByVal Temperature As Double) As Integer
        For i As Integer = 0 To CpData.Count - 1
            CpData(i).Data.Sort(New CpList.DataPointTemperatureComparer)
            If Temperature >= CpData(i).Data(0).Temperature - CpData(i).Data(0).TemperaturePlusMinus Or Temperature <= CpData(i).Data(CpData(i).Data.Count - 1).Temperature + CpData(i).Data(CpData(i).Data.Count - 1).TemperaturePlusMinus Then Return i
        Next
        Return -1
    End Function
    Public Function CpValueFromAllDataPoint(ByVal Temperature As Double) As CpList.DataPoint
        For i As Integer = 0 To CpData.Count - 1
            CpData(i).Data.Sort(New CpList.DataPointTemperatureComparer)
            If Temperature >= CpData(i).Data(0).Temperature - CpData(i).Data(0).TemperaturePlusMinus Or Temperature <= CpData(i).Data(CpData(i).Data.Count - 1).Temperature + CpData(i).Data(CpData(i).Data.Count - 1).TemperaturePlusMinus Then Return CpValueFromDataPoint(i)
        Next
        Return New CpList.DataPoint
    End Function
    Public Function CpValueFromDataPoint(ByVal Temperature As Double, Optional ByVal DataID As Integer = 0) As CpList.DataPoint
        If CpData.Count = 0 Then Return New CpList.DataPoint
        DataID = Math.Min(DataID, CpData.Count - 1)
        With CpData(DataID)
            If .Data.Count = 0 Then Return New CpList.DataPoint
            .Data.Sort(New CpList.DataPointTemperatureComparer)
            For i As Integer = 0 To .Data.Count - 1
                With .Data(i)
                    If .Temperature - .TemperaturePlusMinus <= Temperature And .Temperature + .TemperaturePlusMinus >= Temperature Then
                        Return CpData(0).Data(i)
                    End If
                End With
            Next
            For i As Integer = 0 To CpData(0).Data.Count - 2
                If .Data(i).Temperature - .Data(i).TemperaturePlusMinus <= Temperature And .Data(i + 1).Temperature + .Data(i + 1).TemperaturePlusMinus >= Temperature Then
                    Dim CpValue As Double = (.Data(i + 1).Cp - .Data(i).Cp) / (.Data(i + 1).Temperature - .Data(i).Temperature) * (Temperature - .Data(i).Temperature) + .Data(i).Cp
                    Dim TemperatureUncertainty As Double = Math.Max(.Data(i).TemperaturePlusMinus, .Data(i + 1).TemperaturePlusMinus)
                    Dim CpUncertainty As Double = Math.Max(.Data(i).CpPlusMinus, .Data(i + 1).CpPlusMinus)
                    Return New CpList.DataPoint(Temperature, CpValue, TemperatureUncertainty, CpUncertainty)
                End If
            Next
            Return New CpList.DataPoint
        End With
    End Function

    Public Property CpFitSPLine As SPLine
    Public Property CpFitMin As Double = Double.MaxValue
    Public Property CpFitMax As Double = Double.MinValue
    Public Function InitCpFit() As Boolean
        If CpData Is Nothing Then Return False
        If CpData.Count = 0 Then Return False
        Dim X, Y As New List(Of Double)
        Dim AllData As New CpList
        For i As Integer = 0 To CpData.Count - 1
            AllData.Add(CpData(i))
        Next
        AllData.Data.Sort(New CpList.DataPointTemperatureComparer)
        For i As Integer = 0 To AllData.Data.Count - 1
            X.Add(AllData.Data(i).Temperature)
            Y.Add(AllData.Data(i).Cp)
        Next
        CpFitSPLine = New SPLine()
        CpFitSPLine.Init(X.ToArray, Y.ToArray)

        CpFitMin = AllData.Data(0).Temperature
        CpFitMax = AllData.Data(AllData.Data.Count - 1).Temperature
        Return True
    End Function
    Public Function CpValueFromFitting(ByVal Temperature As Double, Optional ByVal DataID As Integer = 0) As Double
        Dim r As Double = CpValueFromShomateEquation(Temperature)
        If r > 0 Then Return r
        If CpData Is Nothing Then Return -1
        If CpData.Count = 0 Then Return -1
        If CpFitSPLine Is Nothing Then InitCpFit()
        If Temperature < CpFitMin Or Temperature > CpFitMax Then Return -1
        Return CpFitSPLine.Interpolate(Temperature)
    End Function
    Public Sub New()
        Name = ""
        ReactionFormula = ""
        Formula = ""
        ThermoCoefficient = New List(Of ShomateCoeff)
        ThermoDataFileExist = False
        CpData = New List(Of CpList)
    End Sub
    Public Sub New(Optional ByVal Name_ As String = "", Optional ByVal Formula_ As String = "", Optional ByVal ReactionFormula_ As String = "", Optional ByVal ThermoCoeff_ As List(Of ShomateCoeff) = Nothing)
        Name = Name_
        ReactionFormula = ReactionFormula_
        Formula = Formula_
        ThermoDataFileExist = False
        If ThermoCoeff_ IsNot Nothing Then
            ThermoCoefficient = ThermoCoeff_
        Else
            ThermoCoefficient = New List(Of ShomateCoeff)
        End If
        CpData = New List(Of CpList)
    End Sub
End Class
