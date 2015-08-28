<%@ LANGUAGE=VBScript%>

<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body text="#FFFFFF" oncontextmenu=self.event.returnValue=false background="../chatroom/bg1.gif">
<%Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
Response.Flush 
Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0;Data source="
Dim strDatabase, strFolder, strFileName
connstr=Application("yx8_mhjh_connstr")
tempvar=split(connstr,"Data Source=")
temp=  tempvar(1)
temp1=split(temp,"\")
strDatabase=temp1(UBound(temp1))
strFolder=Replace(temp,strDatabase,"")
Private Sub dbCompact()
Dim SourceConn
Dim DestConn
Dim oJetEngine
SourceConn = Jet_Conn_Partial & strFolder & strDatabase
DestConn = Jet_Conn_Partial & strFolder & date() &"压缩备份"& strDatabase
mymdb=strFolder & strDatabase
tomdb=strFolder & date() &"压缩备份"& strDatabase
%>
<table width="80%" border="1" align="center" cellpadding="0" cellspacing="0">
<tr align="center"> 
<td>
<%
If DbExists(DestConn ) Then 
%>
对不起，数据库:<font color=red><%=tomdb%></font>已经存在,使用Ftp删除再试！
<%
Response.End
end if
Set oJetEngine = Server.CreateObject("JRO.JetEngine")
With oJetEngine
	.CompactDatabase SourceConn, DestConn
End With
%> 数据库:<font color=blue><%= mymdb %></font>压缩完成!<br>
      压缩文件文件：<font color=blue><%=tomdb%></font><br>
      请登陆ftp，删除原始数据库：<font color=blue><%=mymdb%></font><br>
      再将压缩数据库：<font color=blue><%=tomdb%></font><br>
      更名成：<font color=blue><%=mymdb%></font>使用！！
<%
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
Select Case Request.form("cmd")
	Case "压缩"
		call dbCompact 
End Select
%> </td>
  </tr>
</table>
<p align="center">&lt;&gt; 用户数据压缩、修复系统 &lt;&gt;</p>
<form method="POST" action="">
<p align="center">
<input type="submit" value="压缩" name="cmd">
</p>
</form>

<p align="center">请选择所压缩的数据库，按压缩开始进行！<br>
  压缩成功后请将原数据库改名，然后将压缩后的数据库改成原数据库！ <b><br>
  </b></p>
