'ChristopherZ
'Spring 2025
'RCET2265
'Accumulate Messages
'https://github.com/Christopher-isu/AccumulateMessages.git

Option Explicit On 'forces all variables to be declared
Option Strict On   'forces all variables to be declared of the correct data type
Imports System

Module MessageProgram
    Sub Main(args As String())
        'uncomment to test interactively
        'Test.Manual()
        Test.Auto()
    End Sub

    Function UserMessages(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static messages As String 'static variables retain their value between function calls

        If clear Then 'clear the messages by isnerting an empty string
            messages = ""
        Else
            If messages <> "" Then  'if messages is not empty, add a new line
                messages = messages & Chr(13) & Chr(10) 'Chr(13) is a carriage return, Chr(10) is a line feed
            End If
            messages = messages & newMessage 'add the new message to the existing messages
        End If

        UserMessages = messages 'return the messages
    End Function


End Module
