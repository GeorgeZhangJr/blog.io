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

Private Sub Form_Load()</br>
WindowsMediaPlayer1.URL = "C:\Users\George\Music\5：15.M4a"</br>
End Sub</br>

然而我文件夾裏有30個對象，這樣衹能操作一餓對象非常麻煩。
然後我加了一個list組建。

Private Sub Form_Load()</br>
List1.Clear</br>
Dim fso, folder, subfolder, file As Object</br></br>
Set fso = CreateObject("scripting.filesystemobject")</br>   </br>                   
Set folder = fso.getfolder("C:\Users\George\Music\")</br></br>
For Each file In folder.Files </br> 
   List1.AddItem file    </br>
Next</br>
Set fso = Nothing</br>
Set folder = Nothing</br>
End Sub</br>

然後也是運行了一下，結果出現前面一段很長的路徑。於是增加一個text1，再把代碼改成下面

Dim path As String</br>
Private Sub Form_Load()</br>
List1.Clear</br>
path = Text1.Text</br>
Dim fso, folder, subfolder, file As Object</br>
Set fso = CreateObject("scripting.filesystemobject")</br>
Set folder = fso.getfolder(path)</br>
For Each file In folder.Files</br>
List1.AddItem Mid(file, Len(path) + 1, Len(file) - Len(path))</br>
Next</br>
Set fso = Nothing</br>
Set folder = Nothing</br>
End Sub</br>

爲了解決播放點擊行數的音樂問題，我特意加了一個變量chosen，添加下列代碼

Private Sub List1_Click()</br>
Dim i As Integer</br>
For i = 0 To List1.ListCount - 1</br>
  If List1.Selected(i) Then</br>
  chosen = List1.List(i)</br>
  End If</br>
Next</br>
WindowsMediaPlayer1.URL = path + chosen</br>
End Sub</br>

測試：雙擊list1可實現功能。

擴展功能：隨機選歌

Private Sub Command2_Click()</br>
chosen = List1.List(Rnd * (List1.ListCount - 1))</br>
WindowsMediaPlayer1.URL = path + chosen</br>

後面的工程我再研究一下明天發
