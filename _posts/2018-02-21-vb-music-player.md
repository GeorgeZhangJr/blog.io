---
layout: post
title: 开源类|VB6.0做的一個音樂播放器
date: 2018-02-21
categories: blog
tags: [开源]
description: 文章金句。
---

首先吐槽一下WIN10的groove播放器的嚴重問題：圖標在任務欄裏每次我想關閉就停了，ipod用的麻煩，所以還是自己做一個。
使用了Windows media player控件，第一次測試代碼如下。

Private Sub Form_Load()
WindowsMediaPlayer1.URL = "C:\Users\George\Music\5：15.M4a"
End Sub

然而我文件夾裏有30個對象，這樣衹能操作一餓對象非常麻煩。
然後我加了一個list組建。

Private Sub Form_Load()
List1.Clear
Dim fso, folder, subfolder, file As Object
Set fso = CreateObject("scripting.filesystemobject")                        
Set folder = fso.getfolder("C:\Users\George\Music\")
For Each file In folder.Files                  
   List1.AddItem file    
Next
Set fso = Nothing
Set folder = Nothing
End Sub

然後也是運行了一下，結果出現前面一段很長的路徑。於是增加一個text1，再把代碼改成下面

Dim path As String
Private Sub Form_Load()
List1.Clear
path = Text1.Text
Dim fso, folder, subfolder, file As Object
Set fso = CreateObject("scripting.filesystemobject")
Set folder = fso.getfolder(path)
For Each file In folder.Files
List1.AddItem Mid(file, Len(path) + 1, Len(file) - Len(path))
Next
Set fso = Nothing
Set folder = Nothing
End Sub

爲了解決播放點擊行數的音樂問題，我特意加了一個變量chosen，添加下列代碼

Private Sub List1_Click()
Dim i As Integer
For i = 0 To List1.ListCount - 1
  If List1.Selected(i) Then
  chosen = List1.List(i)
  End If
Next
WindowsMediaPlayer1.URL = path + chosen
End Sub

測試：雙擊list1可實現功能。

擴展功能：隨機選歌

Private Sub Command2_Click()
chosen = List1.List(Rnd * (List1.ListCount - 1))
WindowsMediaPlayer1.URL = path + chosen

後面的工程我再研究一下明天發
