<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!--#include file="conn.asp"  -->
<%
Dim USERID
Dim strSQL
Dim List,action
Dim name,password,sex,mail,url,icq_img,icqnum,age,gsm,pc,danwei,repassword
Dim years,mons,days,likes,address,photo,doc,mode
Dim ip,counter,n,y,r,s,f,m,sj,times,xingzuo
USERID=Request.Querystring("USERID")
IF Session("username")="" or Session("password")="" or USERID="" Then
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('用户信息丢失，请重新登陆！');" & Chr(13)
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
End if

strSQL="SELECT * FROM List WHERE"
strSQL = strSQL & " name LIKE '%" & Session("username") & "%'"
strSQL = strSQL & " and password LIKE '%" & Session("password") & "%'"
strSQL = strSQL & " and ID LIKE '%" & USERID & "%'"
strSQL =strSQL & " Order by ID DESC"

set list=server.createobject("adodb.recordset")
list.open strSQL,conn,adOpenKeySet,adLockPessimistic
'数据错误处理
if list.eof and list.bof then
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
Response.WRite "错误！没有当前用户！请重新登陆！"
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End
End if

IF  Session("Username")<>list("name") and Session("password")<>list("password") and USERID<>list("id") Then             '如果用户名称等于NULL则证明没有通过验证
Session("username")=""
Session("Password")=""
UserID=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('没有登陆，请重新登陆！');" & Chr(13)
response.write "window.document.location.href='userre.asp';"&Chr(13)
response.write "</script>" & Chr(13)
End if
session("mynameup")=list("name")
Response.Write "OK!通过确认了..."                        '确认用户已经通过确认！
Response.Write Session("Username")

action = Request.Querystring("action")
If action="xiu" Then
Response.Write "修改记录"
name     = trim(Request.Form("name"))      '姓名
password = trim(Request.Form("password"))  '密码
repassword = trim(Request.Form("repassword"))  '密码re
sex      = trim(Request.Form("sex"))       '性别
mail     = trim(Request.Form("mail"))      'E-Mail
URL      = trim(Request.Form("URL"))       '网站链接
icq_img  = trim(Request.Form("icq_img"))   '传呼机图片
icqnum   = trim(Request.Form("icqnum"))    '传呼机号码
AGE      = trim(Request.Form("age"))       '年龄
years    = trim(Request.Form("years"))     '生日的年
mons     = trim(Request.Form("mons"))      '生日的月
days     = trim(Request.Form("days"))      '生日的日
likes    = trim(Request.Form("likes"))     '爱好
gsm      = trim(Request.Form("gsm"))       '手机
pc  	 = trim(Request.Form("pc"))   	   '邮政编码
danwei   = trim(Request.Form("danwei"))    '单位
address  = trim(Request.Form("address"))   '居住地
photo    = trim(Request.Form("photo"))     '相片
doc      = trim(Request.Form("doc"))       '留言
mode     = trim(Request.Form("mode"))      '同学方式
ip       = Request.ServerVariables("REMOTE_ADDR")
if repassword<>password  then _
response.write "密码错误。<a href = javascript:history.back()>返回重输</a>":response.end

'使用ServerVariables读取用户的IP地址，将用户IP放如IP变量中
counter  = "0"
n=Year(date())                                           '取得函数Year年的年份放入n的变量中
y=Month(date())                                          '取的函数Month月的月份放入y的变量中
r=Day(date())                                            '取的函数Day天的天数放入r的变量中
s=Hour(time())                                           '取得函数Hour小时的小时数放入s变量中
f=Minute(time())                                         '取的函数Minute分的分钟数值放入f变量中
m=Second(time())                                         '取的函数Second秒的秒数放入m变量中
if len(y)=1 then y="0" & y                               '如果y的数值等于一位数的话，则y等于0+Y变量，也就是说，把单数前都加一个0，比如01 or 02 or 03
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "年" & y & "月" & r & "日" & " " & weekdayname(weekday(date())) & " " & s & ":" & f & ":" & m
'sj 等于 X年X月X日空格 星期几 几点几分几秒
'WeekDay是提供一个1与7之间的数值，也就是星期几
times    = sj
'#####################星座判断#########################
set list=server.createobject("ADODB.RECORDSET")
list.open strSQL,conn,3,2
list("name")=name
list("password")=password
list("sex")=sex
list("mail")=mail
list("url")=url
list("icqnum")=icqnum
list("icq_img")=icq_img
age=n-years
list("age")=age       '年龄需要根据生日修改，而且需要制作星座代码
list("years")=years
list("mons")=mons
list("days")=days
list("likes")=likes
list("gsm")=gsm
list("pc")=pc
list("danwei")=danwei
list("address")=address
'list("photo")=photo
list("doc")=doc
list("mode")=mode
list("ip")=ip
list("times")=times
list.update
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>个人信息--<%=List("name")%></title><style>
<!--
#tipBox {position: absolute;
width: 160px;
z-index: 100;
border: 1pt black solid;
font-family:宋体;
font-size: 9pt;
background: ffffdf;
visibility: hidden}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
</head>

<body>
<div align="center">
<center>
<table border="1" width="600" cellspacing="0" cellpadding="3" bgcolor="#998CD2" bordercolor="#988CD8">
<tr>
<td width="100%">
<p align="center">您的信息</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" cellspacing="0" cellpadding="3" bordercolor="#988CD8" style="font-family: 宋体; font-size: 9pt" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td width="66" align="center"><a href=main.asp?action=search>⊙-大侠列表</a></td>
<td width="66" align="center"><a href="add.asp">⊙-大侠加入</a></td>
<td width="66" align="center"><a href="userre.asp">⊙-修改资料</a></td>
<td width="66" align="center"><a href="sf.asp?Keys=Login">⊙-超级管理</a></td>
<td width="66" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高级搜索</a></td>
<td width="220" align="center"><font color="#996600">现有</font><font color="#FF6600">
<!--#include file="zongshu.asp" --> </font><font color="#996600">位大侠加入</font></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" cellspacing="0" style="font-family: 宋体; font-size: 9pt" cellpadding="3" bordercolor="#988CD0" height="272">
<tr>
<td width="590" bgcolor="#988CD0" colspan="2" height="12">这是您的修改的结果</td>
</tr>
<!-- <tr>
<td width="590" colspan="2" bgcolor="#EEF1F7" height="12">photo...</td>
</tr> -->
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">姓名：</td>
<td width="433" height="12"><%=list("name")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">门派：</td>
<td width="433" height="12"><%=mode%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">性别：</td>
<td width="433" height="12"><%=list("sex")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">信箱：</td>
<td width="433" height="12"><%=list("mail")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">ICQ/OICQ：</td>
<td width="433" height="12"><img src=<%=list("icq_img")%>  border="0"><%=list("icqnum")%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">生日：</td>
<td width="433" height="12"><%=years%>年<%=mons%>月<%=days%>日</td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">联系电话：</td>
<td width="433" height="12"><%=likes%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="11">邮编/地址：</td>
<td width="433" height="11"><%=address%></td>
</tr>
<tr>
<td width="149" align="right" bgcolor="#CCCCFF" height="12">简短留言：</td>
<td width="433" height="12"><%=doc%></td>
</tr>
<tr>
<td width="582" colspan="2" bgcolor="#CCCCFF" height="25">
<p align="center"><input class="p9" type="submit" value="  返回交友列表  "></p>
</td>
</tr>
</table>
</center>
</div>
</body>

</html>
<%
list.close
response.write "<script language='javascript'>" & chr(13)
response.write "alert('已经成功修改用户资料，请按确定返回！');" & Chr(13)
response.write "window.document.location.href='main.asp?action=search';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End
End if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style>
<!--
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<title>修改资料--<%=List("name")%></title>
</head>

<body>
<form method="POST" action="moduser.asp?USERID=<%=USERID%>&action=xiu">
<div align="center">
<center>
<table border="0" width="600" bgcolor="#988CD0" cellspacing="0" cellpadding="3">
<tr>
<td width="100%">
<p align="center"><b>修 改 资 料</b></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#F8EBD1" bordercolor="#988CD0" cellspacing="0" cellpadding="3" style="font-family: 宋体; font-size: 9pt">
<tr>
<td bgcolor="#F8EBD1" align="center"><a href="main.asp?action=search">⊙-大侠列表</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="add.asp">⊙-大侠加入</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="userre.asp">⊙-修改资料</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="sf.asp?Keys=Login">⊙-超级管理</a></td>
<td bgcolor="#F8EBD1" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高级搜索</a></td>
<td  align="center"><a href style='cursor:hand;'onClick="window.open('upload.asp','up','scrollbars=no,resizable=yes,width=600,height=450')" title='相片上传'>⊙-相片上传</a></td>
</tr>
</table>
<% if isnull(list("photo")) or isempty(list("photo")) then%>
<a href style='cursor:hand;'onClick="window.open('upload.asp','_top','scrollbars=yes,resizable=yes,width=600,height=400')" title='上传相片'>还没有上传相片</a>
<%else %> <font size=2><a href style='cursor:hand;'onClick="javascript:location.reload();" title='刷新'>您的靓照</a><img src="showimg.asp?id=<%=session("Username")%>"><a href style='cursor:hand;'onClick="window.open('upload.asp','up','scrollbars=no,resizable=yes,width=600,height=450')" title='相片重新上传'>我要修改</a>
<%End If%>

</center>
</div>
<div align="center">
<center>
<table border="1" width="600" bgcolor="#DFE1FF" style="font-family: 宋体; font-size: 9pt" cellspacing="0" cellpadding="3" bordercolor="#988CD0">
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">姓名：</td>
<td width="424"><input readonly maxLength="30" name="name" size="30" value="<%=list("name")%>"></td>
<%session("myName")=list("name")%>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">门派：</td>
<td width="424"><input maxLength="30" name="mode" size="30" value="<%=list("mode")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">修改密码：</td>
<td width="424"><input maxLength="30" name="password" size="10" type="password" value="<%=list("Password")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">确认修改密码：</td>
<td width="424"><input maxLength="30" name="repassword" size="10" type="password" value="<%=list("Password")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">性别：</td>
<td width="424"><select class="p9" name="sex" size="1">
<% IF list("sex")="男" Then%>
<option value="">请选择性别</option>
<option selected value="男">男</option>
<option value="女">女</option>
<%ElseIF list("sex")="女" Then%>
<option value="">请选择性别</option>
<option selected value="女">女</option>
<option value="男">男</option>
<%Else%>
<option selected value="">请选择性别</option>
<option value="男">男</option>
<option value="女">女</option>
<%End IF%>
</select></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">信箱：</td>
<td width="424"><input maxLength="50" name="mail" size="30" value="<%=list("mail")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">主页：</td>
<td width="424"><input maxLength="50" name="url" size="30" value="<%=list("URL")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">网络寻呼机：</td>
<td width="424"><select class="p9" name="icq_img" size="1">
<% IF list("icq_img") = "icq.gif" Then%>
<option value="0">没有网络寻呼机</option>
<option selected value="icq.gif">ICQ</option>
<option value="oicq.gif">OICQ</option>
<%ElseIf list("icq_img") = "oicq.gif" Then%>
<option value="0">没有网络寻呼机</option>
<option value="icq.gif">ICQ</option>
<option selected value="oicq.gif">OICQ</option>
<%Else%>
<option selected value="0">没有网络寻呼机</option>
<option value="icq.gif">ICQ</option>
<option value="oicq.gif">OICQ</option>
<%End if%>
</select>号码<input maxLength="10" name="icqnum" size="10" value="<%=list("icqnum")%>"><a href="http://www.icq.com">ICQ下载</a>
<a href="http://www.tencent.com">OICQ下载</a></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">您的生日：</td>
<td width="424"><input maxLength="4" name="years" size="4" value="<%=list("years")%>">年
<input maxLength="2" name="mons" size="2" value="<%=list("mons")%>">月
<input maxLength="2" name="days" size="2" value="<%=list("days")%>">日</td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">联系电话：</td>
<td width="424"><input maxLength="50" name="likes" size="30" value="<%=list("likes")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">手机或呼机：</td>
<td width="424"><input maxLength="50" name="gsm" size="30" value="<%=list("gsm")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">邮政编码：</td>
<td width="424"><input maxLength="6" name="pc" size="30" value="<%=list("pc")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">单位名称：</td>
<td width="424"><input maxLength="50" name="danwei" size="30" value="<%=list("danwei")%>"></td>
</tr>
<tr>
<td width="160" bgcolor="#CCCCFF" align="right">详细地址：</td>
<td width="424"><input maxLength="50" name="address" size="30" value="<%=list("address")%>"></td>
</tr>
<!--    <tr>
<td width="160" bgcolor="#CCCCFF" align="right">相片上传：</td>
<td width="424"><input readonly name="photo" value="<%=list("photo")%>" type="file" disabled >&nbsp;
输入Delete则删除您的相片</td>
</tr> -->
<tr>
<td width="160" bgcolor="#CCCCFF" align="right" valign="top">简短留言：</td>
<td width="424"><textarea cols="39" name="doc" rows="3" wrap="hard" ><%=list("doc")%></textarea></td>
</tr>
<tr>
<td width="584" bgcolor="#CCCCFF" colspan="2">
<p align="center">
<input class="p9" type="submit" value="  修改  ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input class="p9" type="reset" value="  清除  "></td>
</tr>
</table>
</center>
</div>

</forM>

</body>

</html>
