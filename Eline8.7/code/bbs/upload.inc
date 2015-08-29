<%
Dim oUpFileStream

Class Upload_file  
	
	Dim Form,File,Version,Err  
	
	Private Sub Class_Initialize
	 Version = "无组件上传类 Version V1.0"
	 Err = -1
	End Sub
	
	Private Sub Class_Terminate  
	  '清除变量及对像
	  If Err < 0 Then
	    oUpFileStream.Close
	    Form.RemoveAll
	    File.RemoveAll
	    Set Form = Nothing
	    Set File = Nothing
	    Set oUpFileStream = Nothing
	  End If
	End Sub
	   
	Public Sub GetData (RetSize)
	   '定义变量
	  Dim RequestBinData,sStart,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
	  Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
	  Dim iFindStart,iFindEnd
	  Dim iFormStart,iFormEnd,sFormName
	   '代码开始
	  If Request.TotalBytes < 1 Then
	    Err = 1
	    Exit Sub
	  End If
	  If RetSize > 0 Then 
	    If Request.TotalBytes > RetSize Then
	    Err = 2
	    Exit Sub
	    End If
	  End If
	  Set Form = Server.CreateObject ("Scripting.Dictionary")
	  Form.CompareMode = 1
	  Set File = Server.CreateObject ("Scripting.Dictionary")
	  File.CompareMode = 1
	  Set tStream = Server.CreateObject ("ADODB.Stream")
	  Set oUpFileStream = Server.CreateObject ("ADODB.Stream")
	  oUpFileStream.Type = 1
	  oUpFileStream.Mode = 3
	  oUpFileStream.Open 
	  oUpFileStream.Write Request.BinaryRead (Request.TotalBytes)
	  oUpFileStream.Position = 0
	  RequestBinData = oUpFileStream.Read 
	  iFormEnd = oUpFileStream.Size
	  bCrLf = ChrB (13) & ChrB (10)
	  '取得每个项目之间的分隔符
	  sStart = MidB (RequestBinData,1, InStrB (1,RequestBinData,bCrLf)-1)
	  iStart = LenB  (sStart)
	  iFormStart = iStart+2
	  '分解项目
	  Do
	    iInfoEnd = InStrB (iFormStart,RequestBinData,bCrLf & bCrLf)+3
	    tStream.Type = 1
	    tStream.Mode = 3
	    tStream.Open
	    oUpFileStream.Position = iFormStart
	    oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
	    tStream.Position = 0
	    tStream.Type = 2
	    tStream.CharSet = "gb2312"
	    sInfo = tStream.ReadText      
	    '取得表单项目名称
	    iFormStart = InStrB (iInfoEnd,RequestBinData,sStart)-1
	    iFindStart = InStr (22,sInfo,"name=""",1)+6
	    iFindEnd = InStr (iFindStart,sInfo,"""",1)
	    sFormName = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
	    '如果是文件
	    If InStr  (45,sInfo,"filename=""",1) > 0 Then
	      Set oFileInfo = new FileInfo
	      '取得文件属性
	      iFindStart = InStr (iFindEnd,sInfo,"filename=""",1)+10
	      iFindEnd = InStr (iFindStart,sInfo,"""",1)
	      sFileName = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
	      oFileInfo.FileName = GetFileName (sFileName)
	      oFileInfo.FilePath = GetFilePath (sFileName)
	      oFileInfo.FileExt = GetFileExt (sFileName)
	      iFindStart = InStr (iFindEnd,sInfo,"Content-Type: ",1)+14
	      iFindEnd = InStr (iFindStart,sInfo,vbCr)
	      oFileInfo.FileType = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
	      oFileInfo.FileStart = iInfoEnd
	      oFileInfo.FileSize = iFormStart -iInfoEnd -2
	      oFileInfo.FormName = sFormName
	      file.add sFormName,oFileInfo
	    else
	    '如果是表单项目
	      tStream.Close
	      tStream.Type = 1
	      tStream.Mode = 3
	      tStream.Open
	      oUpFileStream.Position = iInfoEnd 
	      oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
	      tStream.Position = 0
	      tStream.Type = 2
	      tStream.CharSet = "gb2312"
	      sFormValue = tStream.ReadText
	      If Form.Exists (sFormName) Then
	        Form (sFormName) = Form (sFormName) & ", " & sFormValue
	        else
	        form.Add sFormName,sFormValue
	      End If
	    End If
	    tStream.Close
	    iFormStart = iFormStart+iStart+2
	    '如果到文件尾了就退出
	  Loop Until  (iFormStart+2) = iFormEnd 
	  RequestBinData = ""
	  Set tStream = Nothing
	End Sub
	
	'取得文件路径
	Private Function GetFilePath (FullPath)
	  If FullPath <> "" Then
	    GetFilePath = Left (FullPath,InStrRev (FullPath, "\"))
	    Else
	    GetFilePath = ""
	  End If
	End Function
	 
	'取得文件名
	Private Function GetFileName (FullPath)
	  If FullPath <> "" Then
	    GetFileName = Mid (FullPath,InStrRev (FullPath, "\")+1)
	    Else
	    GetFileName = ""
	  End If
	End Function
	
	'取得扩展名
	Private Function GetFileExt (FullPath)
	  If FullPath <> "" Then
	    GetFileExt = Mid (FullPath,InStrRev (FullPath, ".")+1)
	    Else
	    GetFileExt = ""
	  End If
	End Function
	
End Class

'文件属性类
Class FileInfo
	
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	
	Private Sub Class_Initialize 
	  FileName = ""
	  FilePath = ""
	  FileSize = 0
	  FileStart = 0
	  FormName = ""
	  FileType = ""
	  FileExt = ""
	End Sub
	  
	'保存文件方法
	Public Function SaveToFile (FullPath)
	  Dim oFileStream,ErrorChar,i
	  SaveToFile = 1
	  If Trim (FullPath) = "" or Right (FullPath,1) = "/" Then exit Function
	  Set oFileStream = CreateObject ("ADODB.Stream")
	  oFileStream.Type = 1
	  oFileStream.Mode = 3
	  oFileStream.Open
	  oUpFileStream.Position = FileStart
	  oUpFileStream.CopyTo oFileStream,FileSize
	  oFileStream.SaveToFile FullPath,2
	  oFileStream.Close
	  Set oFileStream = Nothing 
	  SaveToFile = 0
	End Function
	  
	Public Function GetData
	  oUpFileStream.Position = FileStart
	  GetData = oUpFileStream.Read (FileSize)
  End Function
End Class
%>
