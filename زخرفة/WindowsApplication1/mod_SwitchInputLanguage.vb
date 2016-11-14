''''''''''''''''''''''START''''''''''''''''''>>
'
'Source-Addition by:
'
'              Muhammad Mehmood Iqbal
'               (me_iq_tm@yahoo.Com)
'
'Name:
'
'              mod_SwitchInputLanguage (.vb)
'
'Purpose:
'
'              Language Culture and Keyboard Switch between 40 different languages
'
'Date/Time:
'
'              24th of July 2012, PM
'
'Last Modified by:
'
'              ---------------------
'
'Purpose:
'
'              ---------------------
'
'Date / Time:
'
'              ---------------------
'
''''''''''''''''''''THE END'''''''''''''''''>>

Module mod_SwitchInputLanguage

    'API Declaration
    Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long

    Public Sub SwitchLanguage(ByRef SelectedIndex As Integer)

        'Switch between different language keyboard layouts

        'Select language culture and its related keyboard layout
        Select Case SelectedIndex

            Case 0 'Arabic

                'Load Keyboard
                LoadKeyboardLayout("00000403", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ar-eg")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 1 'Armenian

                'Load Keyboard
                LoadKeyboardLayout("0000042b", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("hy-am")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 2 'Assamese

                'Load Keyboard
                LoadKeyboardLayout("0000044D", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("as-in")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 3 'Divehi

                'Load Keyboard
                LoadKeyboardLayout("00000465", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("dv")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 4 'Dutch

                'Load Keyboard
                LoadKeyboardLayout("00000813", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("nl-be")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 5 'English

                'Load Keyboard
                LoadKeyboardLayout("00000409", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("en-us")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 6 'Hebrew

                'Load Keyboard
                LoadKeyboardLayout("0000040D", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("he")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 7 'Hindi

                'Load Keyboard
                LoadKeyboardLayout("00000439", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("hi-IN")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 8 'German

                'Load Keyboard
                LoadKeyboardLayout("00000407", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("de")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 9 'Gujarati

                'Load Keyboard
                LoadKeyboardLayout("00000447", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("gu")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 10 'Indonesian

                'Load Keyboard
                LoadKeyboardLayout("00000421", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("id")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 11 'Italian

                'Load Keyboard
                LoadKeyboardLayout("00000410", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("it")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 12 'Persian

                'Load Keyboard
                LoadKeyboardLayout("00000429", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("fa")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 13 'Punjabi

                'Load Keyboard
                LoadKeyboardLayout("00000446", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("pa")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 14 'Bangla

                'Load Keyboard
                LoadKeyboardLayout("00000845", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("bn-in")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 15 'Bulgarian

                'Load Keyboard
                LoadKeyboardLayout("00000402", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("bg")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 16 'Bosnian

                'Load Keyboard
                LoadKeyboardLayout("0000641a", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("bs-Cyrl-BA")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 17 'Georgian

                'Load Keyboard
                LoadKeyboardLayout("00000437", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ka")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 18 'Greek

                'Load Keyboard
                LoadKeyboardLayout("00000408", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("el")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 19 'Kannada

                'Load Keyboard
                LoadKeyboardLayout("0000044B", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("kn")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 20 'Lao

                'Load Keyboard
                LoadKeyboardLayout("00000454", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("lo-la")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 21 'Macedonian

                'Load Keyboard
                LoadKeyboardLayout("0000042f", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("mk")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 22 'Malayalam

                'Load Keyboard
                LoadKeyboardLayout("0000044c", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ml-in")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 23 'Malay

                'Load Keyboard
                LoadKeyboardLayout("0000043e", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ms")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 24 'Marathi

                'Load Keyboard
                LoadKeyboardLayout("0000044e", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("mr")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 25 'Mongolian

                'Load Keyboard
                LoadKeyboardLayout("00000450", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("mn")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 26 'Nepali

                'Load Keyboard
                LoadKeyboardLayout("00000461", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ne-np")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 27 'Odia

                'Load Keyboard
                LoadKeyboardLayout("00000448", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("or-in")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 28 'Pashto

                'Load Keyboard
                LoadKeyboardLayout("00000463", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ps-af")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 29 'Russian

                'Load Keyboard
                LoadKeyboardLayout("00000419", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ru")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 30 'Sinhala

                'Load Keyboard
                LoadKeyboardLayout("0000045b", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("si-lk")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 31 'Syriac

                'Load Keyboard
                LoadKeyboardLayout("0000045a", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("syr")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 32 'Tajik 

                'Load Keyboard
                LoadKeyboardLayout("00000428", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("tg-cyrl-tj")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 33 'Tamil

                'Load Keyboard
                LoadKeyboardLayout("00000449", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ta")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 34 'Tatar

                'Load Keyboard
                LoadKeyboardLayout("00000444", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("tt")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 35 'Telugu

                'Load Keyboard
                LoadKeyboardLayout("0000044a", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("te")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 36 'Thai

                'Load Keyboard
                LoadKeyboardLayout("0000041e", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("th")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 37 'Tibetan

                'Load Keyboard
                LoadKeyboardLayout("00000451", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("bo-cn")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)


            Case 38 'Urdu

                'For Urdu Phonetic Keyboard layout, download CRULP Urdu Phonetic Keyboard installer
                'from the following link:
                'http://www.crulp.org/software/localization/keyboards/crulpphonetickbv1.1.html
                'Or download installer file directly
                'http://www.crulp.org/Downloads/localization/keyboards/CRULP_Urdu_Phonetic_kb_v1.1.zip

                'To activate Phonetic Keyboard in textbox, replace following LCID string "00000420"
                'to "a0000420"

                'In some cases, "00000420" also works for Phonetic Keyboard.

                '* LCID stands for Windows Language Code Identifier

                'Load Keyboard
                LoadKeyboardLayout("00000420", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ur")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)


            Case 39 'Uyghur

                'Load Keyboard
                LoadKeyboardLayout("00000480", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("ug-cn")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case 40 'Wolof

                'Load Keyboard
                LoadKeyboardLayout("00000488", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("wo-sn")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)

            Case Else

                'Load Keyboard Default (English)
                LoadKeyboardLayout("00000409", &H2)

                'Change Language Culture
                Dim sys_cul As New System.Globalization.CultureInfo("en-us")
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul)


        End Select

    End Sub

    Public Sub SwitchDefaultLanguage()

        'Set Default language

        'Change Language Culture
        Dim sys_cul_def As New System.Globalization.CultureInfo("en")

        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(sys_cul_def)

    End Sub


    'End of the Module
End Module
