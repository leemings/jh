<!--#Include file="config.asp" -->
<%
'设置密码和帐号

	On Error Resume Next

If Request("name1")="" or Request("password1")="" then
	Response.Write "身份无法确认，停止运行！"
	Response.End
	Else
	Response.Write "验证身份...<br>"
	If Request("name1")=username and Request("password1")=password then
		Response.Write "ok..通过验证！"
		sysop = True
		Session("sysop")= sysop
		call Croom
		Else
		Response.Write "Error!无法通过验证，请检查自己的身份！"
		sysop = False
		Session("sysop") = sysop
		response.End
	End If
Response.End
end if

Sub Croom()
	%>
<HTML>

<HEAD>
<META HTTP-EQUIV="Content-Language" CONTENT="zh-cn">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<META NAME="GENERATOR" CONTENT="Microsoft FrontPage 4.0">
<META NAME="ProgId" CONTENT="FrontPage.Editor.Document">
<TITLE>系统管理</TITLE>
<STYLE>
<!--
body         { font-size: 9pt }
-->
</STYLE>
</HEAD>

<BODY>

<P>请输入要增加的房间名称，每一个房间之间请使用 , 
逗号分割开来，之间不存在空格，头和尾不存在逗号，请不要使用中文的逗号。<BR>
示例：友情聊天室,亲情聊天室,快乐聊天室,英文聊天室,中文聊天室,系统聊天室,某某的家</P>
<P><FONT COLOR="#FF0000">注意：如果不照规定格式输入，将发生系统错误！！</FONT></P>
<FORM METHOD="POST" ACTION="sysop1.asp">
<input type=hidden name='action' value='sysop'>
<P><INPUT TYPE="text" NAME="cr" SIZE="100" STYLE="WIDHT:100% " VALUE="<%=OpenTextFile(Server.MapPath("."), "roomname.asp")%>"><BR>
<INPUT TYPE="submit" VALUE="提交" NAME="B1"><INPUT TYPE="reset" VALUE="全部重写" NAME="B2"></P>
</FORM>

</BODY>

</HTML>

	<%
End Sub

Function OpenTextFile(PathInfo, FileIn)
	On Error Resume Next
	Dim FileObject, InStream
	Set FileObject = Server.CreateObject("Scripting.FileSystemObject")	
	Set InStream = FileObject.OpenTextFile (PathInfo & "\" & FileIn, 1, False, False)	
	OpenTextFile = Instream.ReadALL
	InStream.Close
	Set InStream = Nothing
	If Err Then
		OpenTextFile = "错误号#" & Err.number & VbCrlf
		OpenTextFile = OpenTextFile & Err.description 
		Err.Clear
	End If
End Function
%>