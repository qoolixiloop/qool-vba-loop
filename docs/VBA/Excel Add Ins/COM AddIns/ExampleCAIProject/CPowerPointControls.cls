VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPowerPointControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPowerPointMenus
' This class contains the definitions, creation code, and event procedures for the
' menu items for PowerPoint.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''
' Private reference to the PowerPoint Application.
'''''''''''''''''''''''''''''''''''''''''''''
Private pPowerPointApp As PowerPoint.Application

'''''''''''''''''''''''''''''''''''''''''''''''''
' Reference to the Tools menu in PowerPoint.
'''''''''''''''''''''''''''''''''''''''''''''''''
Private PowerPointToolsMenu As Office.CommandBarPopup

'''''''''''''''''''''''''''''''''''''''''''''
' One test menu item.
'''''''''''''''''''''''''''''''''''''''''''''
Private WithEvents MenuItem1 As Office.CommandBarButton
Attribute MenuItem1.VB_VarHelpID = -1

''''''''''''''''''''''''''''''''''''''''
' All controls created in this module
' should have be stored in this Collection.
' We keep references in a Collection to
' make it easy to delete them when the
' add-in is disconnected. Since we're
' not using the Tag property (to allow
' for separate event handlers for
' each control), we need an easy way
' to get a reference to each control
' when we delete them.
''''''''''''''''''''''''''''''''''''''''
Private ControlsCollection As Collection

'''''''''''''''''''''''''''''''''''''''''
' Property to Get/Set the reference
' to the PowerPoint application.
'''''''''''''''''''''''''''''''''''''''''
Friend Property Get PowerPointApp() As PowerPoint.Application
    Set PowerPointApp = pPowerPointApp
End Property

Friend Property Set PowerPointApp(PPTApp As PowerPoint.Application)
    Set pPowerPointApp = PPTApp
End Property

Friend Sub CreateControls()
'''''''''''''''''''''''''''''''''''''''''''''''''''
' CreateControls
' This is where you would create all of your menu
' items and commandbar controls. Store each one
' in the ControlCollection object so we can
' easily delete them when the add-in disconnects.
' Note that we are NOT using either the Tag
' property or the OnAction property.
'''''''''''''''''''''''''''''''''''''''''''''''''''

Set PowerPointToolsMenu = pPowerPointApp.CommandBars.FindControl(Id:=C_POWERPOINT_TOOLS_MENU_ID)
Set MenuItem1 = PowerPointToolsMenu.Controls.Add(Type:=msoControlButton, temporary:=True)
With MenuItem1
    .Caption = "Click Me PowerPoint"
End With
ControlsCollection.Add Item:=MenuItem1

'''''''''''''''''''''''''
' Repeat the code above
' for each control you
' need to add.
'''''''''''''''''''''''''

End Sub

Friend Sub DeleteControls()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' DeleteControls
' This loops through the ControlsCollection object
' and deletes all of our controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim Ctrl As Office.CommandBarControl
For Ndx = ControlsCollection.Count To 1 Step -1
    Set Ctrl = ControlsCollection(Ndx)
    Ctrl.Delete
    ControlsCollection.Remove Ndx
Next Ndx

End Sub

Private Sub Class_Initialize()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Initialize
' Create a new Collection object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlsCollection = New Collection
End Sub

Private Sub Class_Terminate()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class_Terminate
' Destroy the Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set ControlsCollection = Nothing
End Sub

Private Sub MenuItem1_Click(ByVal Ctrl As Office.CommandBarButton, _
        CancelDefault As Boolean)
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' MenuItem1_Click
' Event procedure for menu item.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    MsgBox Ctrl.Caption
End Sub


