Imports System.IO
Imports System.IO.Compression
Imports System.Windows
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.ApplicationServices



Public Class Form1

    Public Fname As New Object
    Public CSVName As New Object
    Public user_name As String
    Public download_path As String
    Public unzip_path As String



    Public Sub ExcuteMacro(ByVal filename As String, ByVal MacroName As String)
        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oRng As Excel.Range
        oXL = New Excel.Application
        oXL.Visible = False
        oWB = oXL.Workbooks.Open(filename)
        oWB.RunAutoMacros(1)
        oXL.Run(MacroName)
        'MacroName
        oXL.Quit()
        oXL = Nothing
        oWB = Nothing
    End Sub


    'unzip napin painalluksesta
    Private Sub Unzip_b_Click(sender As Object, eventa As EventArgs) Handles Unzip_b.Click

        Dim parts() As String = Split(My.User.Name, "\")
        user_name = parts(1)
        download_path = "C:\Users\" + user_name + "\Downloads\"

        'MsgBox(download_path)

        'Fname on zip tiedosto jota lähdetään avaamaan
        'Dim Fname As New Object

        'luodaan open file dialogi
        Dim OpenFileDialog1 As New OpenFileDialog()
        OpenFileDialog1.Filter = "Zip|*.zip"
        OpenFileDialog1.Title = "Open a zip File"

        'jos painetaan ok nappia
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            'PictureBox1.Load(OpenFileDialog1.FileName)
            Fname = OpenFileDialog1.FileName

        ElseIf OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Exit Sub

        End If

        Dim myfile_name As String
        myfile_name = Fname

        Dim parts_zip() As String = myfile_name.Split(New Char() {"_"c})
        myfile_name = parts_zip(0)

        myfile_name = myfile_name.Replace(".zip", "")
        myfile_name = myfile_name.Replace("D:\", "")
        myfile_name = myfile_name.Replace(download_path, "")
        myfile_name = myfile_name.Replace("C:\", "")

        'MsgBox(myfile_name)

        unzip_path = download_path + myfile_name + "\"


        'vakio extract path
        If Not Directory.Exists(unzip_path) Then
            MkDir(unzip_path)
            MkDir(unzip_path + "PDF")
            MkDir(unzip_path + "DXF")
            If STPBox.Checked = True Then
                MkDir(unzip_path + "STP")
            End If
        End If


        Dim extract_path As String
        extract_path = unzip_path

        Dim pdf_path As String
        pdf_path = unzip_path + "PDF"
        Dim dxf_path As String
        dxf_path = unzip_path + "DXF"
        Dim stp_path As String
        stp_path = unzip_path + "STP"

        'ZipFile.ExtractToDirectory(Fname, "D:\TEMP")

        'luodaan testiluuppia varten muuttujat
        Dim copied_files As New ArrayList
        Dim to_copy As Boolean

        'using käskyllä avataan ja lopuksi suljetaan zip tiedosto Fname
        Using archive As ZipArchive = ZipFile.OpenRead(Fname)
            'lähdetään lukemaan archivea läpi tiedosto tiedostolta
            For Each entry As ZipArchiveEntry In archive.Entries
                'jos archiivista löytyy pdf tiedostotyyppi leimataan to_copy trueksi
                If entry.FullName.EndsWith(".pdf", StringComparison.Ordinal) Then
                    to_copy = True

                    'verrataan onko copied_files listassa jo kyseistä tiedostoa, jos on leimataan to_copy falseksi
                    For Each file_name In copied_files
                        If entry.Name = file_name Then
                            to_copy = False

                        End If
                    Next
                    'jos to_copy on edelleen true lisätään listaan kyseinen tiedostonimi
                    If to_copy = True Then
                        copied_files.Add(entry.Name)

                    End If

                    'luodaan extract kansioon tiedostolle polku

                    'Dim destination_path As String = Path.GetFullPath(Path.Combine(extract_path, entry.Name))
                    Dim destination_path As String = Path.GetFullPath(Path.Combine(pdf_path, entry.Name))

                    'tarkastetaan että extract_path on polussa ja että to_copy on edelleen true leimalla
                    If destination_path.StartsWith(pdf_path, StringComparison.Ordinal) And to_copy = True Then
                        'jos tiedosto sattui olemaan jo kansiossa deleteoidaan se ennen extractointia
                        If File.Exists(destination_path) = True Then
                            File.Delete(destination_path)
                        End If
                        'extraktoidaan tiedosto
                        entry.ExtractToFile(destination_path)
                    End If



                    'jos archiivista löytyy dxf tiedostotyyppi leimataan to_copy trueksi
                ElseIf entry.FullName.EndsWith(".dxf", StringComparison.Ordinal) Then
                    to_copy = True


                    'verrataan onko copied_files listassa jo kyseistä tiedostoa, jos on leimataan to_copy falseksi
                    For Each file_name In copied_files
                        If entry.Name = file_name Then
                            to_copy = False

                        End If
                    Next
                    'jos to_copy on edelleen true lisätään listaan kyseinen tiedostonimi
                    If to_copy = True Then
                        If DXFBox.Checked = True And entry.Length < 1000000 Then
                            copied_files.Add(entry.Name)
                            'MsgBox(entry.Length)
                        End If
                        If DXFBox.Checked = False Then
                            copied_files.Add(entry.Name)
                        End If
                    End If

                    'luodaan extract kansioon tiedostolle polku

                    Dim destination_path As String = Path.GetFullPath(Path.Combine(dxf_path, entry.Name))

                    'tarkastetaan että extract_path on polussa ja että to_copy on edelleen true leimalla
                    If destination_path.StartsWith(dxf_path, StringComparison.Ordinal) And to_copy = True Then
                        'jos tiedosto sattui olemaan jo kansiossa deleteoidaan se ennen extractointia
                        If File.Exists(destination_path) = True Then
                            File.Delete(destination_path)
                        End If
                        'extraktoidaan tiedosto ja testataan SMALL dxf nappula
                        If DXFBox.Checked = True And entry.Length < 1000000 Then
                            entry.ExtractToFile(destination_path)
                            'MsgBox(entry.Length)
                        End If
                        If DXFBox.Checked = False Then
                            entry.ExtractToFile(destination_path)
                        End If

                    End If







                    ''''''''''''''''''''''''''''' 'muut tiedostotyypit juureen

                Else
                        to_copy = True


                    'verrataan onko copied_files listassa jo kyseistä tiedostoa, jos on leimataan to_copy falseksi
                    For Each file_name In copied_files
                        If entry.Name = file_name Then
                            to_copy = False

                        End If
                    Next
                    'jos to_copy on edelleen true lisätään listaan kyseinen tiedostonimi
                    If to_copy = True Then
                        copied_files.Add(entry.Name)
                    End If

                    'luodaan extract kansioon tiedostolle polku

                    Dim destination_path As String = Path.GetFullPath(Path.Combine(extract_path, entry.Name))
                    Dim destination_path_stp As String = Path.GetFullPath(Path.Combine(stp_path, entry.Name))

                    'tarkastetaan että extract_path on polussa ja että to_copy on edelleen true leimalla
                    If destination_path.StartsWith(extract_path, StringComparison.Ordinal) And to_copy = True Then
                        'jos tiedosto sattui olemaan jo kansiossa deleteoidaan se ennen extractointia
                        If File.Exists(destination_path) = True Then
                            File.Delete(destination_path)
                        End If
                        'extraktoidaan tiedosto
                        entry.ExtractToFile(destination_path)

                    End If
                    'jos STEP nappi päällä ja tiedosto pääte on .stp kopiodaan steppi kansioon
                    If destination_path_stp.StartsWith(stp_path, StringComparison.Ordinal) And to_copy = True And STPBox.Checked = True And entry.Name.EndsWith(".stp") Then
                        'jos tiedosto sattui olemaan jo kansiossa deleteoidaan se ennen extractointia
                        If File.Exists(destination_path_stp) = True Then
                            File.Delete(destination_path_stp)
                        End If

                        'poistetaan steppi tiedosto root kansiosta minne kopioitiin edellisessä
                        If File.Exists(destination_path) = True Then
                            File.Delete(destination_path)
                        End If

                        'extraktoidaan tiedosto steppi kansioon
                        entry.ExtractToFile(destination_path_stp)

                    End If

                End If

            Next

        End Using

        MsgBox("Files copied to:" + extract_path)

    End Sub



    Dim objApp As Excel.Application
    Dim objBook_xls As Excel._Workbook
    Dim objBook_csv As Excel._Workbook
    Dim objBook_temp As Excel._Workbook



    Private Sub csv_xls_b_Click(sender As Object, eventa As EventArgs) Handles csv_xls_b.Click



        'MsgBox(unzip_path)

        'luodaan open file dialogi
        Dim OpenFileDialog2 As New OpenFileDialog()
        OpenFileDialog2.Filter = "CSV|*.csv"
        OpenFileDialog2.Title = "Open a CSV File"

        'jos painetaan ok nappia
        If OpenFileDialog2.ShowDialog() = DialogResult.OK Then
            'PictureBox1.Load(OpenFileDialog1.FileName)
            CSVName = OpenFileDialog2.FileName

        ElseIf OpenFileDialog2.ShowDialog() = DialogResult.Cancel Then
            Exit Sub

        End If

        'luodaan uuden tiedoston nimi
        Dim myCSVfile_name As String
        myCSVfile_name = CSVName

        'poistetaan nimestä pois epäoleellinen

        Dim parts1() As String = myCSVfile_name.Split(New Char() {"\"c})
        myCSVfile_name = parts1(UBound(parts1))
        Dim parts2() As String = myCSVfile_name.Split(New Char() {"_"c})
        myCSVfile_name = parts2(0)

        myCSVfile_name = myCSVfile_name.Replace(".csv", "")

        'MsgBox(myCSVfile_name)




        'BOM templaten osoite
        Dim file_to_copy As String
        file_to_copy = File.ReadAllText("C:\Packer_tiedostot\excel_osoite.txt")
        Dim file_to_copy_name As String
        file_to_copy_name = myCSVfile_name + ".xlsm"

        'lukee ajettavat makrot tekstitiedostosta
        Dim macros_to_run_string As String
        macros_to_run_string = File.ReadAllText("C:\Packer_tiedostot\makrot.txt")
        Dim macros() As String = Split(macros_to_run_string, Environment.NewLine)
        'MsgBox(macros(0))

        If Not Len(download_path) > 0 Then
            Dim parts_user() As String = Split(My.User.Name, "\")
            user_name = parts_user(1)
            download_path = "C:\Users\" + user_name + "\Downloads\"
        End If

        Dim file_to_copy_address As String

        If System.IO.File.Exists(file_to_copy) = False Then
            MsgBox("BOM template not found!")
            Exit Sub
        End If



        If Len(unzip_path) > 0 Then
            file_to_copy_address = unzip_path + file_to_copy_name
            'kopioidaan tiedosto unzip kansioon jos kansio on määritelty
            If System.IO.File.Exists(unzip_path + file_to_copy_name) = False Then
                System.IO.File.Copy(file_to_copy, unzip_path + file_to_copy_name)
            Else
                System.IO.File.Delete(unzip_path + file_to_copy_name)
                System.IO.File.Copy(file_to_copy, unzip_path + file_to_copy_name)
            End If

        Else
            file_to_copy_address = download_path + file_to_copy_name
            'kopioidaan tiedosto downloads kansioon jos unzip kansiota ei ole määritelty
            If System.IO.File.Exists(download_path + file_to_copy_name) = False Then
                System.IO.File.Copy(file_to_copy, download_path + file_to_copy_name)
            Else
                System.IO.File.Delete(download_path + file_to_copy_name)
                System.IO.File.Copy(file_to_copy, download_path + file_to_copy_name)
            End If
        End If

        Dim objBooks As Excel.Workbooks
        Dim objSheets_xls As Excel.Sheets
        Dim objSheet_xls1 As Excel._Worksheet
        Dim objSheet_xls2 As Excel._Worksheet
        Dim range_xls As Excel.Range

        Dim objSheets_csv As Excel.Sheets
        Dim objSheet_csv As Excel._Worksheet
        Dim range_csv As Excel.Range
        Dim range_csv_xlsx As Excel.Range

        ' Create a new instance of Excel and start a new workbook.
        objApp = New Excel.Application()
        objBooks = objApp.Workbooks

        'kopioitu bom pohja uudella nimellä
        objBook_xls = objBooks.Open(file_to_copy_address)

        'filereader ennen kuin avaa CSV:n
        Dim fileReader() As String = File.ReadAllLines(CSVName)

        objBook_csv = objBooks.Open(CSVName)
        objSheets_xls = objBook_xls.Worksheets
        objSheet_xls1 = objSheets_xls(1)
        objSheet_xls2 = objSheets_xls(2)

        objSheets_csv = objBook_csv.Worksheets
        objSheet_csv = objSheets_csv(1)


        'MsgBox(fileReader)
        'Dim view_line As String
        'view_line = objSheet_csv.Range("a2").Value

        Dim x As Integer
        x = 1
        For Each view_line In fileReader
            If view_line.StartsWith("View:") Then
                objSheet_csv.Rows(x).Delete(Shift:=Excel.XlDeleteShiftDirection.xlShiftUp)
            End If
            x = x + 1
        Next
        x = Nothing


        'kopioi csv tiedot A:X sarakkeeseen
        objSheet_csv.Range("A:X").Copy()
        objSheet_xls1.Cells(1, 1).PasteSpecial(Excel.XlPasteType.xlPasteValues)

        'MsgBox(objSheet_csv.Range("A2").Value)

        range_xls = objSheet_xls1.Range("A:A")

        Dim test_commas1 As String
        Dim test_commas2 As String
        test_commas1 = fileReader(2)
        test_commas2 = fileReader(3)
        'MsgBox(test_commas1)
        'siirtää kopioidut rivit kolumneihin jos pilkkuerotin
        If test_commas1.StartsWith(",") Or test_commas2.StartsWith(",") Or test_commas1.EndsWith(",") Or test_commas2.EndsWith(",") Then
            'MsgBox("commasep")
            range_xls.TextToColumns(objSheet_xls1.Cells(1, 1), DataType:=Excel.XlTextParsingType.xlDelimited, Comma:=True, TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote, DecimalSeparator:=".", ConsecutiveDelimiter:=False)

        Else
            'range_csv = Nothing
            'objSheet_csv = Nothing
            'objSheets_csv = Nothing
            'objBook_csv.Close()

            'puolipilkulla pitää eritellä solu solulta, koska text to columns ei toimi
            objSheet_csv.Range(objSheet_csv.Cells(1, 1), objSheet_csv.Cells(10000, 30)).Delete()

            Dim start_line As Excel.Range
            start_line = objSheet_csv.Cells(1, 1)

            Dim temp_line As Excel.Range
            temp_line = objSheet_csv.Cells(1, 1)
            Dim i_ As Integer


            For Each line In fileReader
                line = line.Replace("=""", "")
                line = line.Replace("""", "")

                If line.StartsWith("View:") Then

                Else
                    Dim line_parts() As String
                    line_parts = line.Split(";")
                    i_ = 0
                    For Each line_part In line_parts
                        If Len(line_part) > 0 Then
                            'MsgBox(line_part)
                            temp_line.Value = line_part.ToString
                        End If

                        'temp_line.TextToColumns(temp_line, DataType:=Excel.XlTextParsingType.xlDelimited, Semicolon:=True, TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote, DecimalSeparator:=".", ConsecutiveDelimiter:=False)
                        temp_line = temp_line.Offset(0, 1)
                        i_ = i_ + 1
                    Next
                    temp_line = temp_line.Offset(1, -i_)
                End If

            Next

            'kopioi csv tiedot A:X sarakkeeseen

            objSheet_csv.Range(objSheet_csv.Cells(1, 1), objSheet_csv.Cells(10000, 30)).Copy()

            objSheet_xls1.Cells(1, 1).PasteSpecial(Excel.XlPasteType.xlPasteValues)

            temp_line = Nothing
            start_line = Nothing


        End If

        'TÄSSÄ ONGELMAA SAADA TOIMIMAAN UUDEMMILLA EXCELEILLÄ JOTEN PITI KORJATA uudella koodinpätkällä Application.Goto...
        'objSheet_xls2.Activate()
        'objSheet_xls2.Range("A1").Select()
        objSheet_xls2.Application.Goto(objSheet_xls2.Range("A1"), True)

        'AJAA MAKROT, HOX! MAKROT TULEE OLLA KÄYTTÄJÄN NIMEÄMIÄ EXCELISSÄ ETTÄ NIITÄ VOI AJAA! ESIM. MAKRO1 EI KELPAA
        'testaa onko makro nappi päällä
        If MacroBox.Checked = True Then
            Try
                objBook_xls.RunAutoMacros(1)
                'objApp.Run("'" + file_to_copy_address + "'!Macro1")
                For Each line In macros
                    objApp.Run(line)
                Next
                'MsgBox("running macro")
            Catch ex As Exception
                MsgBox("Can't run Macro" & vbCrLf & ex.Message)
            End Try
        Else

        End If


        'datao.SetData(DataFormats.CommaSeparatedValue, datao)
        'Clipboard.SetDataObject(datao, True)
        'objSheet_xls1.Range("A1:D100").Value = Clipboard.GetText

        'Return control of Excel to the user.
        objApp.Visible = True
        objApp.UserControl = True

        'Clean up a little.
        range_xls = Nothing


        objBooks = Nothing

        objSheet_xls1 = Nothing
        objSheet_xls2 = Nothing
        objSheets_xls = Nothing

        'suljetaaan CSV tiedosto
        range_csv = Nothing
        objSheet_csv = Nothing
        objSheets_csv = Nothing




        Me.objBook_xls.Save()
        Me.objBook_xls.Close()
        Me.objBook_csv.Close()


        Me.objApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(objApp)
        objApp = Nothing

        Dim obj1(1) As Process
        obj1 = Process.GetProcessesByName("EXCEL")
        For Each p As Process In obj1
            p.Kill()
        Next

        MsgBox("BOM copied to:" + download_path)

    End Sub



End Class
