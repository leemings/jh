<!--#Include file="config.asp" -->
<%
'����������ʺ�

	On Error Resume Next

If Request("name1")="" or Request("password1")="" then
	Response.Write "����޷�ȷ�ϣ�ֹͣ���У�"
	Response.End
	Else
	Response.Write "��֤���...<br>"
	If Request("name1")=username and Request("password1")=password then
		Response.Write "ok..ͨ����֤��"
		sysop = True
		Session("sysop")= sysop
		call Croom
		Else
		Response.Write "Error!�޷�ͨ����֤�������Լ�����ݣ�"
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
<TITLE>ϵͳ�����һ�������wWw.51eline.com��</TITLE>
<STYLE>
<!--
body         { font-size: 9pt }
-->
</STYLE>
</HEAD>

<BODY>

<P>������Ҫ���ӵķ������ƣ�ÿһ������֮����ʹ�� , 
���ŷָ����֮�䲻���ڿո�ͷ��β�����ڶ��ţ��벻Ҫʹ�����ĵĶ��š�<BR>
ʾ��������������,����������,����������,Ӣ��������,����������,ϵͳ������,ĳĳ�ļ�</P>
<P><FONT COLOR="#FF0000">ע�⣺������չ涨��ʽ���룬������ϵͳ���󣡣�</FONT></P>
<FORM METHOD="POST" ACTION="sysop1.asp">
<input type=hidden name='action' value='sysop'>
<P><INPUT TYPE="text" NAME="cr" SIZE="100" STYLE="WIDHT:100% " VALUE="<%=OpenTextFile(Server.MapPath("."), "roomname.asp")%>"><BR>
<INPUT TYPE="submit" VALUE="�ύ" NAME="B1"><INPUT TYPE="reset" VALUE="ȫ����д" NAME="B2"></P>
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
		OpenTextFile = "�����#" & Err.number & VbCrlf
		OpenTextFile = OpenTextFile & Err.description 
		Err.Clear
	End If
End Function
%>