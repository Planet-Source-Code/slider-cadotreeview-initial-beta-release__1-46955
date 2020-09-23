VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:  GSubclass
' Author:      Steve McMahon (Original) http://www.vbaccelerator.com
'              Graeme Grant (a.k.a. Slider)
' Date:        01/07/2002 (Adapted)
' Version:     01.01.00
' Description: Subclassing wrapper
' Notes:       Adapted to work with the cLODTreeview wrapper DLL/Classes
'              -=-=-= Warning! =-=-=-
'              Subclassing can crash the IDE if stopped at the wrong point
'              in execution.
'
'===========================================================================

Option Explicit

Public Enum EMsgResponse
    emrConsume      '## Process instead of original WindowProc
    emrPostProcess  '## Process after original WindowProc
    emrPreprocess   '## Process before original WindowProc
End Enum

Public Sub AttachMessage(iwp As cLODTreeview, ByVal hwnd As Long, ByVal iMsg As Long)
    MSubclass.AttachMessage iwp, hwnd, iMsg
End Sub

Public Sub DetachMessage(iwp As cLODTreeview, ByVal hwnd As Long, ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
    MSubclass.DetachMessage iwp, hwnd, iMsg
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
   CurrentMessage = MSubclass.CurrentMessage
End Property

Public Function CallOldWindowProc(ByVal hwnd As Long, _
                                  ByVal iMsg As Long, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."

   CallOldWindowProc = MSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)

End Function
