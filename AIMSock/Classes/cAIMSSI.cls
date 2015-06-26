VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAIMSSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function RequestSSIRights() As cAIMPacketBuffer
    Set RequestSSIRights = New cAIMPacketBuffer
    RequestSSIRights.InsertSNAC ServerStoredInformation, SSISIDs.CS_RequestRights, 0, 0, &H130002
End Function

Public Function RequestSSIData() As cAIMPacketBuffer
    Set RequestSSIData = New cAIMPacketBuffer
    RequestSSIData.InsertSNAC ServerStoredInformation, SSISIDs.CS_RequestSSI, 0, 0, &H130005
    RequestSSIData.InsertDWORD &H12345678   ' Time since last local update
    RequestSSIData.InsertWORD &H165         ' Number of local items
End Function

Public Function ActivateSSI() As cAIMPacketBuffer
    Set ActivateSSI = New cAIMPacketBuffer
    ActivateSSI.InsertSNAC ServerStoredInformation, SSISIDs.CS_ActivateSSI, 0, 0, &H130007
End Function

Public Function AddSSIItem(ByVal ScreenName As String, ByVal GroupID As Long, ByVal BuddyID As Long) As cAIMPacketBuffer
    Set AddSSIItem = New cAIMPacketBuffer
    
    With AddSSIItem
        .InsertSNAC ServerStoredInformation, SSISIDs.CS_AddItem, 0, 0, &H130008
        .InsertString ScreenName
        .InsertWORD GroupID
        .InsertWORD BuddyID
        .InsertWORD SSIItem.Buddy
        .InsertWORD 0   ' No TLVs
    End With
End Function

Public Function ModifySSIGroup(ByVal GroupName As String, ByVal GroupID As Long, ByVal MemberBuddies As String) As cAIMPacketBuffer
    Set ModifySSIGroup = New cAIMPacketBuffer
    
    With ModifySSIGroup
        .InsertSNAC ServerStoredInformation, SSISIDs.CS_ModifyGroup, 0, 0, &H130009
        .InsertString GroupName
        .InsertWORD GroupID
        .InsertWORD 0
        .InsertWORD SSIItem.Group
        .InsertTLV &HC8, MemberBuddies
    End With
End Function

Public Function RemoveSSIItem(ByVal strScreenName As String, ByVal GroupID As Long, ByVal BuddyID As Long) As cAIMPacketBuffer
    Set RemoveSSIItem = New cAIMPacketBuffer
    
    With RemoveSSIItem
        .InsertSNAC ServerStoredInformation, SSISIDs.CS_RemoveItem, 0, 0, &H13000A
        .InsertString strScreenName
        .InsertWORD GroupID
        .InsertWORD BuddyID
        .InsertWORD SSIItem.Buddy
        .InsertWORD 0
    End With
End Function

Public Function SendPreModify() As cAIMPacketBuffer
    Set SendPreModify = New cAIMPacketBuffer
    SendPreModify.InsertSNAC ServerStoredInformation, SSISIDs.CS_BeginMod, 0, 0, &H130011
End Function

Public Function SendPostModify() As cAIMPacketBuffer
    Set SendPostModify = New cAIMPacketBuffer
    SendPostModify.InsertSNAC ServerStoredInformation, SSISIDs.CS_EndMod, 0, 0, &H130012
End Function
