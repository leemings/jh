<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
nickname=Session("aqjh_name")
grade=Int(Session("aqjh_grade"))
if grade<>10 then Response.Redirect "manerr.asp?id=100"
Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0;Data source="
Dim strDatabase, strFolder, strFileName
xp="aqjh.asp"
dataxp="../aqjh_data/"&xp
strFolder = Server.MapPath(dataxp)
strFolder = Replace(strFolder,xp,"")
strDatabase =xp
Private Sub dbCompact()
Dim SourceConn
Dim DestConn
Dim oJetEngine
SourceConn = Jet_Conn_Partial & strFolder & strDatabase
DestConn = Jet_Conn_Partial & strFolder & date() &"压缩备份"& strDatabase
mymdb=strFolder & strDatabase
tomdb=strFolder & date() &"压缩备份"& strDatabase
If DbExists(DestConn ) Then 
	Response.Write ("对不起，数据库:"& tomdb &"已经存在,使用Ftp删除再试！!") 
	Response.End
end if
Set oJetEngine = Server.CreateObject("JRO.JetEngine")
With oJetEngine
	.CompactDatabase SourceConn, DestConn
End With
Response.Write ("数据库:<font color=blue>" & mymdb & "</font>压缩完成!<br>压缩文件文件：<font color=blue>"&tomdb&"</font><br><br>请登陆ftp，删除原始数据库：<font color=blue>"&mymdb&"</font><br>再将压缩数据库：<font color=blue>"&tomdb&"</font><br>更名成：<font color=blue>"&mymdb&"</font>使用！！")
End Sub
Public function DbExists(byVal dbPath) 
'查找数据库文件是否存在 
On Error resume Next 
Dim c 
Set c = Server.CreateObject("ADODB.Connection") 
c.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath 
If Err.number<>0 Then 
Err.Clear 
DbExists = false 
else 
DbExists = True 
End If 
set c = nothing 
End function 

%>
<%
' Compact database and tell the user the database is optimized
Select Case Request.form("cmd")
	Case "压缩"
		call dbCompact 
End Select
%>
<LINK href=css/css.css type=text/css rel=stylesheet>
<title><%=Application("aqjh_chatroomname")%>维护程序</title>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><b><font color=blue>用户数据压缩、修复系统</b></p>
<form method="POST" action="">
<p align="center">
<input type="submit" value="压缩" name="cmd">
</p>
</form>

<p align="center"><font color="#000000">请选择所压缩的数据库，按压缩开始进行！</font><br>
  此压缩数据库操作<font color="#0000FF"><b>免Fso系统 </b></font>需要使用Ftp上传相配合完成！<br>
  <font color="#FF0000"><b>强列要求备份原始数据库文件，否则一切后果自负！<br>
  </b></font><br>执行此操作需要将IE安全设置成低才可以执行！</p>