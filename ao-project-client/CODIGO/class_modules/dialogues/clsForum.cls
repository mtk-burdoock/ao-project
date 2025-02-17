VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_MENSAJES_FORO As Byte = 30
Private Const MAX_ANUNCIOS_FORO As Byte = 5
Private CantPosts(0 To 2) As Byte
Private CantStickys(0 To 2) As Byte

Public Sub AddPost(ByVal ForumType As eForumType, ByRef Title As String, ByRef Author As String, ByRef MsgPost As String, ByVal Sticky As Boolean)
    With Foros(ForumType)
        If Sticky Then
            If CantStickys(ForumType) < MAX_ANUNCIOS_FORO Then _
                CantStickys(ForumType) = CantStickys(ForumType) + 1
            Call MoveArray(ForumType, Sticky)
            .StickyTitle(1) = Title
            .StickyAuthor(1) = Author
            .StickyPost(1) = MsgPost
        Else
            If CantPosts(ForumType) < MAX_MENSAJES_FORO Then _
                CantPosts(ForumType) = CantPosts(ForumType) + 1
            Call MoveArray(ForumType, Sticky)
            .GeneralTitle(1) = Title
            .GeneralAuthor(1) = Author
            .GeneralPost(1) = MsgPost
        End If
    End With
End Sub

Public Sub ClearForums()
    Dim Index As Long
    If CantPosts(eForumType.ieGeneral) > 0 Then
        For Index = 1 To CantPosts(eForumType.ieGeneral)
            Foros(eForumType.ieGeneral).GeneralAuthor(Index) = vbNullString
            Foros(eForumType.ieGeneral).GeneralTitle(Index) = vbNullString
            Foros(eForumType.ieGeneral).GeneralPost(Index) = vbNullString
        Next Index
        CantPosts(eForumType.ieGeneral) = 0
    End If
    If CantPosts(eForumType.ieREAL) > 0 Then
        For Index = 1 To CantPosts(eForumType.ieREAL)
            Foros(eForumType.ieREAL).GeneralAuthor(Index) = vbNullString
            Foros(eForumType.ieREAL).GeneralTitle(Index) = vbNullString
            Foros(eForumType.ieREAL).GeneralPost(Index) = vbNullString
        Next Index
        CantPosts(eForumType.ieREAL) = 0
    End If
    If CantPosts(eForumType.ieCAOS) > 0 Then
        For Index = 1 To CantPosts(eForumType.ieCAOS)
            Foros(eForumType.ieCAOS).GeneralAuthor(Index) = vbNullString
            Foros(eForumType.ieCAOS).GeneralTitle(Index) = vbNullString
            Foros(eForumType.ieCAOS).GeneralPost(Index) = vbNullString
        Next Index
        CantPosts(eForumType.ieCAOS) = 0
    End If
    If CantStickys(eForumType.ieGeneral) > 0 Then
        For Index = 1 To CantStickys(eForumType.ieGeneral)
            Foros(eForumType.ieGeneral).StickyAuthor(Index) = vbNullString
            Foros(eForumType.ieGeneral).StickyTitle(Index) = vbNullString
            Foros(eForumType.ieGeneral).StickyPost(Index) = vbNullString
        Next Index
        CantStickys(eForumType.ieGeneral) = 0
    End If
    If CantStickys(eForumType.ieREAL) > 0 Then
        For Index = 1 To CantStickys(eForumType.ieREAL)
            Foros(eForumType.ieREAL).StickyAuthor(Index) = vbNullString
            Foros(eForumType.ieREAL).StickyTitle(Index) = vbNullString
            Foros(eForumType.ieREAL).StickyPost(Index) = vbNullString
        Next Index
        CantStickys(eForumType.ieREAL) = 0
    End If
    If CantStickys(eForumType.ieCAOS) > 0 Then
        For Index = 1 To CantStickys(eForumType.ieCAOS)
            Foros(eForumType.ieCAOS).StickyAuthor(Index) = vbNullString
            Foros(eForumType.ieCAOS).StickyTitle(Index) = vbNullString
            Foros(eForumType.ieCAOS).StickyPost(Index) = vbNullString
        Next Index
        CantStickys(eForumType.ieCAOS) = 0
    End If
End Sub

Public Function GetNroPost(ByVal ForumType As eForumType) As Byte
    GetNroPost = CantPosts(ForumType)
End Function

Public Function GetNroSticky(ByVal ForumType As eForumType) As Byte
    GetNroSticky = CantStickys(ForumType)
End Function

Private Sub MoveArray(ByVal ForumType As eForumType, ByVal Sticky As Boolean)
    Dim i As Long
    With Foros(ForumType)
        If Sticky Then
            Dim Count_stickys As Long
                Count_stickys = CantStickys(ForumType)
            For i = Count_stickys To 2 Step -1
                .StickyTitle(i) = .StickyTitle(i - 1)
                .StickyPost(i) = .StickyPost(i - 1)
                .StickyAuthor(i) = .StickyAuthor(i - 1)
            Next i
        Else
            Dim Count_posts As Long
                Count_posts = CantPosts(ForumType)
            For i = Count_posts To 2 Step -1
                .GeneralTitle(i) = .GeneralTitle(i - 1)
                .GeneralPost(i) = .GeneralPost(i - 1)
                .GeneralAuthor(i) = .GeneralAuthor(i - 1)
            Next i
        End If
    End With
End Sub
