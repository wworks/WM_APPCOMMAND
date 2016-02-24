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
                     handle, New IntPtr(CommandCode))

    End Sub

    Public Class Browser
        Private Enum BrowserValues
            APPCOMMAND_BROWSER_BACKWARD = &H10000
            APPCOMMAND_BROWSER_FAVORITES = &H60000
            APPCOMMAND_BROWSER_FORWARD = &H20000
            APPCOMMAND_BROWSER_HOME = &H70000
            APPCOMMAND_BROWSER_REFRESH = &H30000
            APPCOMMAND_BROWSER_SEARCH = &H50000
            APPCOMMAND_BROWSER_STOP = &H40000
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
        Public Shared Sub StopBrowser(Handle As Integer)
            DoCommand(Handle, BrowserValues.APPCOMMAND_BROWSER_STOP)
        End Sub
    End Class
    Public Class Media
        Public Enum MediaValues
            APPCOMMAND_MEDIA_CHANNEL_DOWN = &H34000
            APPCOMMAND_MEDIA_CHANNEL_UP = &H33000
            APPCOMMAND_MEDIA_FAST_FORWARD = &H31000
            APPCOMMAND_MEDIA_NEXTTRACK = &HB000
            APPCOMMAND_MEDIA_PAUSE = &H2F000
            APPCOMMAND_MEDIA_PLAY = &H2E000
            APPCOMMAND_MEDIA_PLAY_PAUSE = &HE0000
            APPCOMMAND_MEDIA_PREVIOUSTRACK = &HC0000
            APPCOMMAND_MEDIA_RECORD = &H30000
            APPCOMMAND_MEDIA_REWIND = &H32000
            APPCOMMAND_MEDIA_STOP = &HD0000
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
            DoCommand(Handle, MediaValues.APPCOMMAND_MEDIA_PLAY_PAUSE)
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
            APPCOMMAND_MIC_ON_OFF_TOGGLE = &H2C000
            APPCOMMAND_MICROPHONE_VOLUME_DOWN = &H19000
            APPCOMMAND_MICROPHONE_VOLUME_MUTE = &H18000
            APPCOMMAND_MICROPHONE_VOLUME_UP = &H1A000
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
                APPCOMMAND_VOLUME_DOWN = &H90000
                APPCOMMAND_VOLUME_MUTE = &H80000
                APPCOMMAND_VOLUME_UP = &HA0000
            End Enum
            Public Shared Sub Down(Handle As Integer)
                DoCommand(Handle, VolumeValues.APPCOMMAND_VOLUME_DOWN)
            End Sub
            Public Shared Sub Up(Handle As Integer)
                DoCommand(Handle, VolumeValues.APPCOMMAND_VOLUME_UP)
            End Sub
            Public Shared Sub Mute(Handle As Integer)
                DoCommand(Handle, VolumeValues.APPCOMMAND_VOLUME_MUTE)
            End Sub
        End Class
        Public Class Bass
            Public Enum BassValues
                APPCOMMAND_BASS_BOOST = &H14000
                APPCOMMAND_BASS_DOWN = &H13000
                APPCOMMAND_BASS_UP = &H15000

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
                APPCOMMAND_TREBLE_DOWN = &H16000
                APPCOMMAND_TREBLE_UP = &H17000
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
            APPCOMMAND_LAUNCH_MAIL = &HF0000
            APPCOMMAND_LAUNCH_MEDIA_SELECT = &H10000
            APPCOMMAND_LAUNCH_APP1 = &H11000
            APPCOMMAND_LAUNCH_APP2 = &H12000
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
                APPCOMMAND_COPY = &H24000
                APPCOMMAND_CUT = &H25000
                APPCOMMAND_PASTE = &H26000
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
                APPCOMMAND_REPLY_TO_MAIL = &H27000
                APPCOMMAND_FORWARD_MAIL = &H28000
                APPCOMMAND_SEND_MAIL = &H29000
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
            APPCOMMAND_FIND = &H1C000
            APPCOMMAND_HELP = &H1B000
            APPCOMMAND_NEW = &H1D000
            APPCOMMAND_OPEN = &H1E000
            APPCOMMAND_CLOSE = &H1F000
            APPCOMMAND_SAVE = &H20000
            APPCOMMAND_PRINT = &H21000
            APPCOMMAND_UNDO = &H22000
            APPCOMMAND_REDO = &H23000

            APPCOMMAND_SPELL_CHECK = &H2A000
            APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE = &H2B000
            APPCOMMAND_CORRECTION_LIST = &H2D000
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
    Private Enum AllCommandCodes
        APPCOMMAND_BASS_BOOST = &H14000
        APPCOMMAND_BASS_DOWN = &H13000
        APPCOMMAND_BASS_UP = &H15000
        APPCOMMAND_BROWSER_BACKWARD = &H10000
        APPCOMMAND_BROWSER_FAVORITES = &H60000
        APPCOMMAND_BROWSER_FORWARD = &H20000
        APPCOMMAND_BROWSER_HOME = &H70000
        APPCOMMAND_BROWSER_REFRESH = &H30000
        APPCOMMAND_BROWSER_SEARCH = &H50000
        APPCOMMAND_BROWSER_STOP = &H40000
        APPCOMMAND_CLOSE = &H1F000
        APPCOMMAND_COPY = &H24000
        APPCOMMAND_CORRECTION_LIST = &H2D000
        APPCOMMAND_CUT = &H25000
        APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE = &H2B000
        APPCOMMAND_FIND = &H1C000
        APPCOMMAND_FORWARD_MAIL = &H28000
        APPCOMMAND_HELP = &H1B000
        APPCOMMAND_LAUNCH_APP1 = &H11000
        APPCOMMAND_LAUNCH_APP2 = &H12000
        APPCOMMAND_LAUNCH_MAIL = &HF0000
        APPCOMMAND_LAUNCH_MEDIA_SELECT = &H10000
        APPCOMMAND_MEDIA_CHANNEL_DOWN = &H34000
        APPCOMMAND_MEDIA_CHANNEL_UP = &H33000
        APPCOMMAND_MEDIA_FAST_FORWARD = &H31000
        APPCOMMAND_MEDIA_NEXTTRACK = &HB000
        APPCOMMAND_MEDIA_PAUSE = &H2F000
        APPCOMMAND_MEDIA_PLAY = &H2E000
        APPCOMMAND_MEDIA_PLAY_PAUSE = &HE0000
        APPCOMMAND_MEDIA_PREVIOUSTRACK = &HC0000
        APPCOMMAND_MEDIA_RECORD = &H30000
        APPCOMMAND_MEDIA_REWIND = &H32000
        APPCOMMAND_MEDIA_STOP = &HD0000
        APPCOMMAND_MICROPHONE_VOLUME_DOWN = &H19000
        APPCOMMAND_MICROPHONE_VOLUME_MUTE = &H18000
        APPCOMMAND_MICROPHONE_VOLUME_UP = &H1A000
        APPCOMMAND_MIC_ON_OFF_TOGGLE = &H2C000
        APPCOMMAND_NEW = &H1D000
        APPCOMMAND_OPEN = &H1E000
        APPCOMMAND_PASTE = &H26000
        APPCOMMAND_PRINT = &H21000
        APPCOMMAND_REDO = &H23000
        APPCOMMAND_REPLY_TO_MAIL = &H27000
        APPCOMMAND_SAVE = &H20000
        APPCOMMAND_SEND_MAIL = &H29000
        APPCOMMAND_SPELL_CHECK = &H2A000
        APPCOMMAND_TREBLE_DOWN = &H16000
        APPCOMMAND_TREBLE_UP = &H17000
        APPCOMMAND_UNDO = &H22000
        APPCOMMAND_VOLUME_DOWN = &H90000
        APPCOMMAND_VOLUME_MUTE = &H80000
        APPCOMMAND_VOLUME_UP = &HA0000
    End Enum
End Class
