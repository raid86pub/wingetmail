Imports System.Threading.Thread
Imports System.IO

Module rdmail
        Const IDX_GREP = 0
        Const IDX_HOST = 1
        Const IDX_USER = 2
        Const IDX_PSWD = 3
        Dim gMs() As String = {"RE-VBCMD", "pop.sina.com", "rdstk01@sina.com", "Lrf@12345"}

        '  The mailman object is used for receiving (POP3)
        '  and sending (SMTP) email.
        Dim mailman As New Chilkat.MailMan()

        Sub Main(ByVal xArgs() As String)
                Dim oNow As Date = Now()

                Today = "2015/12/28"
                Debug.Print("hack: " & Now().ToString())
                '  Any string argument automatically begins the 30-day trial.
                If (mailman.UnlockComponent("Anything for 30-day trial") <> True) Then
                        Debug.Print("Component unlock failed: " & Now().ToString())
                        Exit Sub
                End If
                Today = oNow.ToString("yyyy/MM/dd")
                Debug.Print("hack: " & Now().ToString())
                '  TimeOfDay = "12:22:26"
                '  Debug.Print(Now().ToString())

                For i = 0 To xArgs.Length - 1
                        gMs(i) = xArgs(i)
                Next

                DeleteMail()
        End Sub

        Sub DeleteMail()
                Dim success As Boolean
                Dim aCells() As String

                '  Set the POP3 server's hostname
                mailman.MailHost = gMs(IDX_HOST)

                '  Set the POP3 login/password.
                mailman.PopUsername = gMs(IDX_USER)

                If (File.Exists("RDTOKEN.TXT")) Then
                        mailman.PopPassword = Trim(File.ReadAllText("RDTOKEN.TXT"))
                Else
                        mailman.PopPassword = gMs(IDX_PSWD)
                End If

                '  Set the ImmediateDelete property = False
                '  When the DeleteEmail method is called, the POP3 session
                '  will remain open and the emails will be deleted all at once
                '  when the session closes.
                mailman.ImmediateDelete = False

                Dim bundle As Chilkat.EmailBundle

                '  Download email headers.
                bundle = mailman.GetAllHeaders(1)

                If (bundle Is Nothing) Then
                        Debug.Print(mailman.LastErrorText)
                        Exit Sub
                End If


                Dim i As Integer
                Dim email As Chilkat.Email
                Dim aKeys() As String
                Debug.Print("bundle.MessageCount = " & bundle.MessageCount)
                For i = 0 To bundle.MessageCount - 1
                        email = bundle.GetEmail(i)
                        Debug.Print(email.Subject)
                        aKeys = gMs(IDX_GREP).Split("+")
                        For j = 0 To aKeys.Count - 1
                                If (aKeys(j).Length < 2) Then
                                        Continue For
                                End If
                                If (email.Subject.IndexOf(aKeys(j)) >= 0) Then
                                        '  If you decide the email is to be deleted, call DeleteEmail.
                                        '  This marks the email for deletion on the POP3 server.
                                        success = mailman.DeleteEmail(email)
                                        If (success <> True) Then
                                                Debug.Print(mailman.LastErrorText)
                                                Exit Sub
                                        End If

                                        aCells = Trim(email.Subject).Split()
                                        If (aCells.Count > 0) Then
                                                Console.WriteLine(Trim(email.FromAddress) & " " & aCells(aCells.Count - 1))
                                        End If

                                        Exit For
                                End If
                        Next
                Next

                '  Make sure the POP3 session is ended to finalize the deletes.
                success = mailman.Pop3EndSession()
                If (success <> True) Then
                        Debug.Print(mailman.LastErrorText)
                End If
        End Sub
End Module
