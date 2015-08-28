<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp" -->
<%
dim copyright
dim Aff
Dim strSQL
Dim List,IDS,counter
IDS=Request("ID")
session("mynameup")=""
IF IDS="" Then
REsponse.Write "错误！丢失信息！"
Response.End
End If
Aff = Chr(34)

'set list=server.createobject("adodb.recordset")
'list.Open "select * from list where id="+cint(Request.Querystring("IDS")) ,conn,1
'list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'数据错误处理
'counter=list("counter")
'counter=counter + 1


'strSQL="SELECT * FROM List ID="+cint(Request.Querystring("IDS"))
'strSQL =strSQL & " Order by ID"
'set list=server.createobject("ADODB.RECORDSET")
'list.open strSQL,conn,3,2

strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  ID    LIKE '"&request("id")&"'"
strSQL =strSQL & " Order by ID DESC"

set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
if list.eof and list.bof then
Response.WRite "错误！没有当前用户！"
Response.End
End If

'if list.eof and list.bof then
' Response.WRite "错误！没有当前用户！"
' Response.End
'End if
counter=list("counter")
counter=counter+1
list("counter")=counter
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>个人信息--<%=List("name")%></title><STYLE type=text/css>BODY {
FONT-SIZE: 12px; FONT-FAMILY: "宋体","Arial Narrow","Times New Roman"
}
P {
FONT-SIZE: 12px; COLOR: #82668a; FONT-FAMILY: "宋体","Arial Narrow","Times New Roman"
}
TD {
FONT-SIZE: 12px; COLOR: #82668a; FONT-FAMILY: "宋体","Arial Narrow","Times New Roman"
}
A:link {
FONT-SIZE: 12px; COLOR: #978de9; FONT-FAMILY: 宋体; TEXT-DECORATION: none
}
A:visited {
FONT-SIZE: 12px; COLOR: #978de9; TEXT-DECORATION: none
}
P {
FONT-SIZE: 12px; FONT-FAMILY: 宋体
}
DIV {
FONT-SIZE: 12px; FONT-FAMILY: 宋体
}
A:hover {
COLOR: red
}
</STYLE>
<script LANGUAGE="JavaScript">
bg = new Array(33);
bg[0] = '../riji/000.gif'; bg[1] = '../riji/001.gif'; bg[2] = '../riji/002.gif'; bg[3] = '../riji/003.gif'; bg[4] = '../riji/004.gif'; bg[5] = '../riji/005.gif'; bg[6] = '../riji/006.gif'; bg[7] = '../riji/007.gif'; bg[8] = '../riji/008.gif'; bg[9] = '../riji/009.gif'; bg[10] = '../riji/010.gif'; bg[11] = '../riji/011.gif'; bg[12] = '../riji/012.gif'; bg[13] = '../riji/013.gif'; bg[14] = '../riji/014.gif'; bg[15] = '../riji/015.gif'; bg[16] = '../riji/016.gif'; bg[17] = '../riji/017.gif'; bg[18] = '../riji/018.gif'; bg[19] = '../riji/019.gif'; bg[20] = '../riji/010.gif'; bg[21] = '../riji/021.gif'; bg[22] = '../riji/022.gif'; bg[23] = '../riji/023.gif'; bg[24] = '../riji/024.gif'; bg[25] = '../riji/025.gif'; bg[26] = '../riji/026.gif'; bg[27] = '../riji/027.gif'; bg[28] = '../riji/028.gif'; bg[29] = '../riji/029.gif'; bg[30] = '../riji/030.gif'; bg[31] = '../riji/031.gif'; bg[32] = '../riji/032.gif';
index = Math.floor(Math.random() * bg.length); document.write("<BODY BACKGROUND="+bg[index]+">"); </script> </head>

<body>
<form method="POST" action="main.asp" >
<div align="center">
<input name=action type=hidden value=search>
<center>
<table border="0" width="600" cellspacing="0" cellpadding="3">
<tr>
<td width="100%">
<p align="center"><%=List("name")%>的注册信息</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" cellspacing="0" cellpadding="3" bordercolor="#988CD8" style="font-family: 宋体; font-size: 9pt" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="center"><a href=javascript:history.back();>⊙-返回前页</a></td>
<td align="center"><a href=main.asp?action=search>⊙-所有大侠</a></td>
<td align="center"><a href="add.asp">⊙-大侠加入</a></td>
<td align="center"><a href="userre.asp"></a><a href="userre.asp">⊙-修改资料</a></td>
<td align="center"><a href="sf.asp?Keys=Login">⊙-超级管理</a></td>
<td align="center"><a href="admsearch.asp?keys=admsearch">⊙-高级搜索</a></td>
<td align="center"><font color="#996600">现有</font><font color="#FF6600">
<!--#include file="zongshu.asp" -->
</font><font color="#996600">位大侠加入</font></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" cellspacing="0" style="font-family: 宋体; font-size: 9pt" cellpadding="3" bordercolor="#988CD0" height="272">
<tr>
<td bgcolor="#988CD0" colspan="6" height="12">这是<%=List("name")%>的登陆情况</td>
</tr>
<tr>
<td  colspan="6" bgcolor="#EEF1F7" height="12" align=center>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
<a href style='cursor:hand;'onClick="window.open('iwantup.asp','_top','scrollbars=yes,resizable=yes,width=600,height=330')" title='上传相片'>还没有上传相片</a>
<%else %> <img src="showimg.asp?id=<%=List("name")%>">
<%End If%>
</td>
</tr>

<tr>
<td width=60 align="right" bgcolor="#CCCCFF" height="12">姓名：</td>
<td height="12"><font color=0000ff size=5 face=隶书,黑体><%=List("name")%></font></td>
<td width=80 align="right" bgcolor="#CCCCFF" height="12">门派：</td>
<td height="12"><%=list("mode")%></td>
<td width=40 align="right" bgcolor="#CCCCFF" height="12">性别：</td>
<td height="12"><%=list("sex")%></td>
</tr>
<tr>
<td align="right" bgcolor="#CCCCFF" height="12">信箱：</td>
<td height="12"> <a href=mailto:<%=list("mail")%>><%=list("mail")%>
</td>
<td align="right" bgcolor="#CCCCFF" height="12">主页：</td>
<td height="12" colspan="3">        <%If list("url")<>"" Then%>
<a href=<%=list("URL")%>  target="_blank"><%=list("urL")%><%End If%> </a> </td>
</tr>

<tr>
<td align="right" bgcolor="#CCCCFF" height="12">ICQ/OICQ：</td>
<td height="12"> <%If list("icq_img")<>"" and list("icqnum")<>"" Then
If list("icq_img")="icq.gif" Then
%><img src=icq.gif>ICQ:<%=list("icqnum")%> <%Else%> <img src=oicq.gif>OICQ:<%=list("icqnum")%>
<%End If%> <%End IF%> </td>
<td align="right" bgcolor="#CCCCFF" height="12">生日：</td>
<td height="12" colspan="3"><%=list("years")%>年<%=list("mons")%>月<%=list("days")%>日</td>
</tr>
<tr>
<td align="right" bgcolor="#CCCCFF" height="12">联系电话：</td>
<td height="12"><%=list("likes")%></td>
<td align="right" bgcolor="#CCCCFF" height="12">手机或呼机：</td>
<td height="12" colspan="3"><%=list("gsm")%></td>
</tr>
<tr>
<td align="right" bgcolor="#CCCCFF" height="11">邮政编码：</td>
<td height="11"><%=list("pc")%></td>
<td align="right" bgcolor="#CCCCFF" height="11">单位名称：</td>
<td height="11" colspan="3"><%=list("danwei")%></td>
</tr>
<tr>
<td align="right" bgcolor="#CCCCFF" height="11">详细地址：</td>
<td height="11" colspan="5"><%=list("address")%></td>
</tr>
<tr>
<td align="right" bgcolor="#CCCCFF" height="12">简短留言：</td>
<td height="12" colspan="5"><%=list("doc")%></td>
</tr>
<tr>
<td  colspan="6" bgcolor="#CCCCFF" height="25">
<p align="center">
<input class="p9" type="submit" value="  返回交友列表  ">
</p>
</td>
</tr>
</table>
</center>
</div>
</form>
</body>

</html>
<%List.CLose%>
<%Response.End%>
