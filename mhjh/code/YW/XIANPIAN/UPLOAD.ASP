<%
if Session("mynameup")="" then
%>
<script language="vbscript">
MsgBox "◎要上传相片，请从[修改资料]处进入！！"
location.href = "userre.asp"
</script>
<%
end if
%>

<!--#include file="conn.asp"-->
<% rs=conn.execute("select * from List where name='"&session("myName")&"'")%>
<HTML><HEAD><TITLE><%=session("myName")%>_照片上载</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
<!--td {  font-family: 宋体; font-size: 9pt}
body {  font-family: 宋体; font-size: 9pt}
select {  font-family: 宋体; font-size: 9pt}
A {text-decoration: none; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #666666; font-family: "宋体"; font-size: 9pt} .txt {  font-family: "宋体"; font-size: 10.5pt}
--></style>

</HEAD>
<body bgcolor="#8d8051" text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0">
<a href=javascript:window.close();>关闭本窗</a> | <a href=javascript:location.reload();>刷新显示</a>
<TABLE width=100% align="center">
<TBODY>
<TR>
<TD align=middle width="704" height="294">
<div align="right"><BR>
</div>
<P align="center"><%=session("myName")%>_照片上载</a>
<form name="Form" enctype="multipart/form-data" action="upphoto.asp" method=post>
<TABLE border=1 cellPadding=0 cellSpacing=1 width="443" bordercolor="#000000" align="center">
<TBODY>
<tr>
<TD bgColor=#b2a265 vAlign=center>
<P>在你的硬盘上要上传的图片文件: <INPUT name=mefile type=file>
<P align="center">
<input type="submit" name="Submit" value="上传照片">
<input type="reset" name="cancel" value="哎呀错了">
<br>
注意：只允许上传文件格式<font color="#00FFFF">.jpg</font>或<font color="#00FFFF">.gif</font>的照片</P>
</TD>
</TR>
<TR>
<TD bgColor=#abcdef>&nbsp;<br> <center><% if isnull(rs("photo")) or isempty(rs("photo")) then%>
<img src="photo.jpg" width="160" height="221"> <%else %> <img src="showimg.asp?id=<%=session("myName")%>">
<%end if%> <br>
<font color="#000000"> 这是你的照片吗?</font> </TD></tr>          </TBODY>
</TABLE>
</FORM>
</TD>
</TR>
</TBODY>
</TABLE>
<center><a href=javascript:window.close();>关闭本窗</a> | <a href=javascript:location.reload();>刷新显示</a>
</BODY></HTML>
