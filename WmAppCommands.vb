Imports System.Runtime.InteropServices

Public Class WmAppCommands

    Public Const WM_APPCOMMAND As Integer = &H319

    <DllImport("user32.dll")>
    Public Shared Function SendMessageW(ByVal hWnd As IntPtr,
               ByVal Msg As Integer, ByVal wParam As IntPtr,
               ByVal lParam As IntPtr) As IntPtr
    End Function
    Public Shared Sub DoCommand(handle As Integer, CommandCode As Integer)
        SendMessageW(handle, WM_APPCOMMAND,
                     handle, New IntPtr(CommandCode << 16))

    End Sub

    Public Class Browser
        Private Enum BrowserValues
            APPCOMMAND_BROWSER_BACKWARD = 1
            APPCOMMAND_BROWSER_FAVORITES = 6
            APPCOMMAND_BROWSER_FORWARD = 2
            APPCOMMAND_BROWSER_HOME = 7
            APPCOMMAND_BROWSER_REFRESH = 3
            APPCOMMAND_BROWSER_SEARCH = 5
            APPCOMMAND_BROWSER_STOP = 4
        End Enum

        Public Shared Sub Backward(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_BACKWARD)
        End Sub
        Public Shared Sub Forward(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_FORWARD)
        End Sub
        Public Shared Sub Favorites(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_FAVORITES)
        End Sub
        Public Shared Sub Home(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_HOME)
        End Sub
        Public Shared Sub Refresh(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_REFRESH)
        End Sub
        Public Shared Sub Search(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_SEARCH)
        End Sub
        Public Shared Sub StopDownload(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_STOP)
        End Sub
    End Class
    Public Class Media
        Public Enum MediaValues
            APPCOMMAND_MEDIA_CHANNEL_DOWN = 52
            APPCOMMAND_MEDIA_CHANNEL_UP = 51
            APPCOMMAND_MEDIA_FAST_FORWARD = 49
            APPCOMMAND_MEDIA_NEXTTRACK = 11
            APPCOMMAND_MEDIA_PAUSE = 47
            APPCOMMAND_MEDIA_PLAY = 46
            APPCOMMAND_MEDIA_PLAY_PAUSE = 14
            APPCOMMAND_MEDIA_PREVIOUSTRACK = 12
            APPCOMMAND_MEDIA_RECORD = 48
            APPCOMMAND_MEDIA_REWIND = 50
            APPCOMMAND_MEDIA_STOP = 13
        End Enum
        Public Shared Sub ChannelDown(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_CHANNEL_DOWN)
        End Sub
        Public Shared Sub ChannelUp(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_CHANNEL_UP)
        End Sub
        Public Shared Sub FastForward(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_FAST_FORWARD)
        End Sub
        Public Shared Sub NextTrack(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_NEXTTRACK)
        End Sub
        Public Shared Sub Pause(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_PAUSE)
        End Sub
        Public Shared Sub Play(Handle As Integer)

            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_PLAY)
        End Sub
        Public Shared Sub PlayPause(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_PLAY_PAUSE << 16)
        End Sub
        Public Shared Sub PreviousTrack(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_PREVIOUSTRACK)
        End Sub
        Public Shared Sub Record(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_RECORD)
        End Sub
        Public Shared Sub Rewind(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_REWIND)
        End Sub
        Public Shared Sub StopMedia(Handle As Integer)
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_STOP)
        End Sub
    End Class
    Public Class Microphone
        Public Enum MicrophoneValues
            APPCOMMAND_MIC_ON_OFF_TOGGLE = 44
            APPCOMMAND_MICROPHONE_VOLUME_DOWN = 25
            APPCOMMAND_MICROPHONE_VOLUME_MUTE = 24
            APPCOMMAND_MICROPHONE_VOLUME_UP = 26
        End Enum
        Public Shared Sub OnOffToggle(Handle As Integer)
            DoCommand(Handle, MicrophoneValues.APPCOMMAND_MIC_ON_OFF_TOGGLE)
        End Sub
        Public Shared Sub VolumeDown(Handle As Integer)
            DoCommand(Handle, MicrophoneValues.APPCOMMAND_MICROPHONE_VOLUME_DOWN)
        End Sub
        Public Shared Sub VolumeUp(Handle As Integer)
            DoCommand(Handle, MicrophoneValues.APPCOMMAND_MICROPHONE_VOLUME_UP)
        End Sub
        Public Shared Sub VolumeMute(Handle As Integer)
            DoCommand(Handle, MicrophoneValues.APPCOMMAND_MICROPHONE_VOLUME_MUTE)
        End Sub
    End Class

    Public Class Sound
        Public Class Volume
            Public Enum VolumeValues
                APPCOMMAND_VOLUME_DOWN = 9
                APPCOMMAND_VOLUME_MUTE = 8
                APPCOMMAND_VOLUME_UP = 10
            End Enum
            Public Shared Sub Down(Handle As Integer)
                DoCommand(Handle, VolumeValues.APPCOMMAND_VOLUME_DOWN << 16)
            End Sub
            Public Shared Sub Up(Handle As Integer)
                DoCommand(Handle, VolumeValues.APPCOMMAND_VOLUME_UP)
            End Sub
            Public Shared Sub Mute(Handle As Integer)
                DoCommand(Handle, VolumeValues.APPCOMMAND_VOLUME_MUTE << 16)
                DoCommand(Handle, 8 << 16)
            End Sub
        End Class
        Public Class Bass
            Public Enum BassValues
                APPCOMMAND_BASS_BOOST = 20
                APPCOMMAND_BASS_DOWN = 19
                APPCOMMAND_BASS_UP = 21

            End Enum
            Public Shared Sub Down(Handle As Integer)
                DoCommand(Handle, BassValues.APPCOMMAND_BASS_DOWN)
            End Sub
            Public Shared Sub Up(Handle As Integer)
                DoCommand(Handle, BassValues.APPCOMMAND_BASS_UP)
            End Sub
            Public Shared Sub Boost(Handle As Integer)
                DoCommand(Handle, BassValues.APPCOMMAND_BASS_BOOST)
            End Sub
        End Class
        Public Class Treble
            Public Enum TrebleValues
                APPCOMMAND_TREBLE_DOWN = 22
                APPCOMMAND_TREBLE_UP = 23
            End Enum
            Public Shared Sub Down(Handle As Integer)
                DoCommand(Handle, TrebleValues.APPCOMMAND_TREBLE_DOWN)
            End Sub
            Public Shared Sub Up(Handle As Integer)
                DoCommand(Handle, TrebleValues.APPCOMMAND_TREBLE_UP)
            End Sub
        End Class
    End Class
    Public Class Launch
        Public Enum LaunchValues
            APPCOMMAND_LAUNCH_MAIL = 15 '&HF0000
            APPCOMMAND_LAUNCH_MEDIA_SELECT = 16
            APPCOMMAND_LAUNCH_APP1 = 17
            APPCOMMAND_LAUNCH_APP2 = 18
        End Enum
        Public Shared Sub Mail(Handle As Integer)
            DoCommand(Handle, LaunchValues.APPCOMMAND_LAUNCH_MAIL)
        End Sub
        Public Shared Sub MediaSelect(Handle As Integer)
            DoCommand(Handle, LaunchValues.APPCOMMAND_LAUNCH_MEDIA_SELECT)
        End Sub
        Public Shared Sub App1(Handle As Integer)
            DoCommand(Handle, LaunchValues.APPCOMMAND_LAUNCH_APP1)
        End Sub
        Public Shared Sub App2(Handle As Integer)
            DoCommand(Handle, LaunchValues.APPCOMMAND_LAUNCH_APP2)
        End Sub


    End Class
    Public Class Actions
        Public Class Clipboard
            Public Enum ClipboardValues
                APPCOMMAND_COPY = 36
                APPCOMMAND_CUT = 37
                APPCOMMAND_PASTE = 38
            End Enum
            Public Shared Sub CopySelection(Handle As Integer)
                DoCommand(Handle, ClipboardValues.APPCOMMAND_COPY)
            End Sub
            Public Shared Sub CutSelection(Handle As Integer)
                DoCommand(Handle, ClipboardValues.APPCOMMAND_CUT)
            End Sub
            Public Shared Sub Paste(Handle As Integer)
                DoCommand(Handle, ClipboardValues.APPCOMMAND_PASTE)
            End Sub
        End Class
        Public Class Mail
            Public Enum MailValues
                APPCOMMAND_REPLY_TO_MAIL = 39
                APPCOMMAND_FORWARD_MAIL = 40
                APPCOMMAND_SEND_MAIL = 41
            End Enum
            Public Shared Sub ReplyToMail(Handle As Integer)
                DoCommand(Handle, MailValues.APPCOMMAND_REPLY_TO_MAIL)
            End Sub
            Public Shared Sub ForwardMail(Handle As Integer)
                DoCommand(Handle, MailValues.APPCOMMAND_FORWARD_MAIL)
            End Sub
            Public Shared Sub SendMail(Handle As Integer)
                DoCommand(Handle, MailValues.APPCOMMAND_SEND_MAIL)
            End Sub
        End Class

        Public Enum ActionsValues
            APPCOMMAND_FIND = 28
            APPCOMMAND_HELP = 27
            APPCOMMAND_NEW = 29
            APPCOMMAND_OPEN = 30
            APPCOMMAND_CLOSE = 31
            APPCOMMAND_SAVE = 32
            APPCOMMAND_PRINT = 33
            APPCOMMAND_UNDO = 34
            APPCOMMAND_REDO = 35

            APPCOMMAND_SPELL_CHECK = 42
            APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE = 43
            APPCOMMAND_CORRECTION_LIST = 45
        End Enum
        Public Shared Sub Find(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_FIND)
        End Sub
        Public Shared Sub Help(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_HELP)
        End Sub
        Public Shared Sub Neww(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_NEW)
        End Sub
        Public Shared Sub Open(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_OPEN)
        End Sub
        Public Shared Sub Close(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_CLOSE)
        End Sub
        Public Shared Sub Save(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_SAVE)
        End Sub
        Public Shared Sub Print(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_PRINT)
        End Sub
        Public Shared Sub Undo(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_UNDO)
        End Sub
        Public Shared Sub Redo(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_REDO)
        End Sub
        Public Shared Sub SpellCheck(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_SPELL_CHECK)
        End Sub
        Public Shared Sub DictateOrCommandControlToggle(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE)
        End Sub
        Public Shared Sub CorrectionList(Handle As Integer)
            DoCommand(Handle, ActionsValues.APPCOMMAND_CORRECTION_LIST)
        End Sub
    End Class

End Class
