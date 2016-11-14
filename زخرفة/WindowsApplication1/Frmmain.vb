Option Compare Text
Imports Helpers
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO

''Done On 28/09/2012 4:00:00 PM

Public Class Frmmain

    Private WithEvents kbHook As New KeyboardHook
    Dim strin As String = Nothing

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.Setting.Clear()
        Dim TXA As New List(Of String)

        For i = 1 To 33
            TXA.Add(GetObja(i, True).Text)
        Next

        My.Settings.Setting.AddRange(TXA.ToArray)
        My.Settings.Save()
        End
    End Sub

    Function ArabicFT(ByVal KeyT As String) As [String]
        Dim TmpArabicFT As [String] = ""
        Dim DataAr(), DataAt() As [String]
        On Error Resume Next
        Dim Txt1, Txt2, Txt3, Txt4, Txt5, Txt6, Txt7, Txt8, Txt9, Txt10, Txt11, Txt12, Txt13, Txt14, Txt15, Txt16, Txt17, Txt18, Txt19, Txt20, Txt21, Txt22, Txt23, Txt24, Txt25, Txt26, Txt27, Txt28, Txt29, Txt30, Txt31, Txt32, Txt33 As [String]
        Txt1 = " " + TextBox1.Text
        Txt2 = " " + TextBox2.Text
        Txt3 = " " + TextBox3.Text
        Txt4 = " " + TextBox4.Text
        Txt5 = " " + TextBox5.Text
        Txt6 = " " + TextBox6.Text
        Txt7 = " " + TextBox7.Text
        Txt8 = " " + TextBox8.Text
        Txt9 = " " + TextBox9.Text
        Txt10 = " " + TextBox10.Text
        Txt11 = " " + TextBox11.Text
        Txt12 = " " + TextBox12.Text
        Txt13 = " " + TextBox13.Text
        Txt14 = " " + TextBox14.Text
        Txt15 = " " + TextBox15.Text
        Txt16 = " " + TextBox16.Text
        Txt17 = " " + TextBox17.Text
        Txt18 = " " + TextBox18.Text
        Txt19 = " " + TextBox19.Text
        Txt20 = " " + TextBox20.Text
        Txt21 = " " + TextBox21.Text
        Txt22 = " " + TextBox22.Text
        Txt23 = " " + TextBox23.Text
        Txt24 = " " + TextBox24.Text
        Txt25 = " " + TextBox25.Text
        Txt26 = " " + TextBox26.Text
        Txt27 = " " + TextBox27.Text
        Txt28 = " " + TextBox28.Text
        Txt29 = " " + TextBox29.Text
        Txt30 = " " + TextBox30.Text
        Txt31 = " " + TextBox31.Text
        Txt32 = " " + TextBox32.Text
        Txt33 = " " + TextBox33.Text
        DataAr = Split(Txt1 + Txt2 + Txt3 + Txt4 + Txt5 + Txt6 + Txt7 + Txt8 + Txt9 + Txt10 + Txt11 + Txt12 + Txt13 + Txt14 + Txt15 + Txt16 + Txt17 + Txt18 + Txt19 + Txt20 + Txt21 + Txt22 + Txt23 + Txt24 + Txt25 + Txt26 + Txt27 + Txt28 + Txt29 + Txt30 + Txt31 + Txt32 + Txt33 + " ", " ")
        DataAt = Split(" ا" + " ب" + " ت" + " ث" + " ج" + " ح" + " خ" + " د" + " ذ" + " ر" + " ز" + " س" + " ش" + " ص" + " ض" + " ط" + " ظ" + " ع" + " غ" + " ف" + " ق" + " ك" + " ل" + " م" + " ن" + " لا" + " ء" + " ه" + " ؤ" + " ئ" + " و" + " ى" + " ي", " ")
        For i = 1 To UBound(DataAr)
            If DataAt(i) = KeyT Then
                TmpArabicFT = DataAr(i)
                Exit For
            End If
        Next i
        'If TmpArabicFT = "" Then TmpArabicFT = KeyT
        Return TmpArabicFT
    End Function

    Private Function ConvertByteArrayToString(ByVal byteArray As Byte()) As String
        Dim enc As Encoding = Encoding.Unicode
        Dim text As String = enc.GetString(byteArray)
        Return text
    End Function


    Function ArabicMUsoftware(ByVal KeyT As String) As String
        KeyT = Replace(KeyT, "Oem6", "]", , , CompareMethod.Text)
        KeyT = Replace(KeyT, "Oem4", "[", , , CompareMethod.Text)
        KeyT = Replace(KeyT, "OemQuotes", "'", , , CompareMethod.Text)
        KeyT = Replace(KeyT, "Oem2", "/", , , CompareMethod.Text)
        KeyT = Replace(KeyT, "OemComma", ",", , , CompareMethod.Text)
        KeyT = Replace(KeyT, "OemSemiColon", ";", , , CompareMethod.Text)
        KeyT = Replace(KeyT, "OemPeriod", ".", , , CompareMethod.Text)


        Dim Eng() As String = Split("q w e r t y u i o p [ ] a s d f g h j k l ; ' z x c v b n m , . /", " ")
        Dim Ara() As String = Split("ض ص ث ق ف غ ع ه خ ح ج د ش س ي ب ل ا ت ن م ك ط ئ ء ؤ ر لا ى ة و ز ظ", " ")

        Dim charpos As Integer

        For charpos = 0 To UBound(Eng)
            If Eng(charpos) = KeyT Then
                Return Ara(charpos)
            End If
        Next
        For charpos = 0 To UBound(Ara)
            If KeyT.Contains(Ara(charpos)) Then
                Return KeyT
            End If
        Next
        Return KeyT
    End Function


    Function TypeArabic(ByVal TextFT As String) As String
        Dim DataAr() As String = {}
        Dim SetData As String = ""
        For i = 1 To Len(TextFT)
            SetData = SetData & ArabicFT(Mid$(TextFT, i, 1))
        Next i
        TypeArabic = SetData
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With My.Computer.Clipboard
            .Clear()
            Application.DoEvents()
            .SetText(Text2.Text)
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Text2.Text = TypeArabic(Text1.Text)
    End Sub

    Private Function GetObja(Num As Integer, Bola As Boolean) As Control
        For Each I As Control In XXX.Controls
            If Bola Then
                If TypeOf I Is TextBox Then
                    If CInt(I.Name.Replace("TextBox", "")) = Num Then
                        Return I
                    End If
                End If
            End If
            If Bola = False Then
                If TypeOf I Is Label Then
                    If CInt(I.Name.Replace("C", "")) = Num Then
                        Return I
                    End If
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Sub Frmmain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If My.Settings.Setting Is Nothing Then
            My.Settings.Setting = New System.Collections.Specialized.StringCollection
            For i = 1 To 33
                GetObja(i, True).Text = GetObja(i, False).Text
            Next
        Else
            My.Settings.Reload()
            For i = 1 To 33
                GetObja(i, True).Text = My.Settings.Setting(i - 1)
            Next
        End If
        SwitchLanguage(0)

    End Sub

    <DllImport("user32")> _
    Private Shared Function GetKeyboardLayoutName(ByVal sb As System.Text.StringBuilder) As Integer
    End Function
    Dim lk As [String]
    Property Las As [String]
        Get
            Return lk
        End Get
        Set(value As [String])
            lk = value

            Timer1.Enabled = True
        End Set
    End Property

    'Public Sub DisableInput(ByVal makeDisabled As Boolean)
    '    Dim j As New Project1.Class1
    '    ToolStripStatusLabel6.Text = j.DisableInput(makeDisabled)
    '    'Dim n As Boolean = 'BlockInput(makeDisabled)
    '    'ToolStripStatusLabel6.Text = n
    'End Sub

    Private Sub kbHook_KeyDown(Key As Keys) Handles kbHook.KeyDown
        Dim LastPressed As String




        LastPressed = [Enum].GetName(GetType(Keys), Key)

        LastPressed = LastPressed.Replace("NumPad", "")

        Dim sb As New StringBuilder(" "c, 256)
        Dim len As Integer
        len = GetKeyboardLayoutName(sb)
        ToolStripStatusLabel2.Text = LastPressed
        If sb.ToString() = "00000401" Then
            'ToolStripStatusLabel6.Text = (DisableInput(1))

            'Threading.Thread.Sleep(100)
            Dim TmpKey As [String] = ArabicMUsoftware(LastPressed)
            If TmpKey <> "" Then

                If CheckBox1.Checked Then
                    '.Replace(TmpKey, "")

                    If ArabicFT(TmpKey).Length <> 0 Then
                        'TextBox1
                        blockinput(True)
                        'SendKeys.Send(Las)
                        Las = ArabicFT(TmpKey)
                        Timer1.Enabled = True
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Las IsNot Nothing AndAlso (Las <> "Enter" And Las <> "Space") Then
            If CheckBox1.Checked Then
                '   ToolStripStatusLabel6.Text = BlockInput(False)
                blockinput(False)
                SendKeys.Send(Las)
                Las = ""

            End If
        End If
        Timer1.Enabled = False
    End Sub


    Private Sub kbHook_KeyUp(Key As Keys) Handles kbHook.KeyUp
        Dim LastPressed As String
        LastPressed = [Enum].GetName(GetType(Keys), Key)

        LastPressed = LastPressed.Replace("NumPad", "")

        Dim sb As New StringBuilder(" "c, 256)
        Dim len As Integer
        len = GetKeyboardLayoutName(sb)

        ToolStripStatusLabel4.Text = LastPressed
        If sb.ToString() = "00000401" Then
            Dim TmpKey As String = ArabicMUsoftware(LastPressed)
            If TmpKey <> "" Then
                ToolStripStatusLabel4.Text = (TmpKey)
                ' BlockInput(False)
                blockinput(False)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For i = 1 To 33
            GetObja(i, True).Text = GetObja(i, False).Text
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        FrmTable.ShowDialog()
    End Sub

    Private Sub Text1_ChangeUICues(sender As Object, e As UICuesEventArgs) Handles Text1.ChangeUICues

    End Sub

    Private Sub Text1_Click(sender As Object, e As EventArgs) Handles Text1.Click
        SwitchLanguage(0)
    End Sub

    Private Sub Text1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text1.KeyPress
        '      SendKeys.Send("ﭪ")
    End Sub
    
    'Private Structure M2

    'End Structure
   
    <Serializable()> Private Structure M
        Public Contains As Byte()
    End Structure

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim TXA() As M
        ReDim TXA(32)

        For i = 1 To 33
            TXA(i - 1).Contains = (Encoding.Unicode.GetBytes(GetObja(i, True).Text))
        Next

        'FileOpen(1, IO.Path.GetDirectoryName(Application.ExecutablePath) & "\Backup.db", OpenMode.Random)
        'FilePut(1, TXA)
        'FileClose(1)
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim fStream As New FileStream(IO.Path.GetDirectoryName(Application.ExecutablePath) & "\Backup.db", FileMode.OpenOrCreate)

        bf.Serialize(fStream, TXA) ' write to file
      
        'My.Computer.FileSystem.WriteAllBytes(IO.Path.GetDirectoryName(Application.ExecutablePath) & "\Backup.db" _
        '                                     , TXA.ToArray, False)
        fStream.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Dim TXA As [String]()
        If IO.File.Exists(IO.Path.GetDirectoryName(Application.ExecutablePath) & "\Backup.db") Then
            'FileOpen(1, IO.Path.GetDirectoryName(Application.ExecutablePath) & "\Backup.db", OpenMode.Input)
            'FileGetObject(1, TXA)
            'FileClose(1)
            Dim TXA() As M
            Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
            Dim fStream As New FileStream(IO.Path.GetDirectoryName(Application.ExecutablePath) & "\Backup.db", FileMode.Open)

            TXA = bf.Deserialize(fStream)


            fStream.Close()
            For i = 1 To 33
                GetObja(i, True).Text = ConvertByteArrayToString(TXA(i - 1).Contains)
            Next

        End If

    End Sub

    'Private Function ConvertByteArrayToString(ByVal byteArray As Byte()) As String
    '    Dim enc As Encoding = Encoding.Unicode
    '    Dim text As String = enc.GetString(byteArray)
    '    Return text
    'End Function

  
    Private Sub Text1_TextChanged(sender As Object, e As EventArgs) Handles Text1.TextChanged

    End Sub
End Class
