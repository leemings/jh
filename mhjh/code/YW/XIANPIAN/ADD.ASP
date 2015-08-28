
<%
uname=session("yx8_mhjh_username")

if uname="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 where 姓名='" &uname& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是剑侠中人或者连接超时"
conn.close
response.end
mycorp=rs("门派")
end if%>
<%if request("action")<>"add" then%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<style type="text/css">
<!--
table {  font-size: 9pt}
td {  font-size: 9pt}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>大侠加入</title></head>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" topmargin="0" leftmargin="0">

<div align="center"><center><form method="POST" action="add.asp"><table border="1" width="600" height="310" cellspacing="0" bgcolor="#DCE2FF" bordercolor="#988CD0">

<tr><td width="594" colspan="2" bgcolor="#FFFFFF">
<p align="center"><font face="宋体" size="3"><b>请大侠详细填写你的资料</b></font>
</td></tr>
<tr><td width="594" colspan="2">

<div align="center"><center><table border="1" width="100%" bordercolorlight="#000000" cellspacing="0" cellpadding="4" bordercolordark="#FFFFFF" bgcolor="#F8EBD1">
<tr>
<td width="13%" align="center"><a href=main.asp?action=search>⊙-大侠列表</a></td>
<td width="13%" align="center">⊙大侠加入</td>
<td width="13%" align="center"><a href=userre.asp>⊙-修改资料</a></td>
<td width="13%" align="center"><a href="sf.asp?Keys=Login">⊙-超级管理</a></td>
<td width="13%" align="center"><a href="admsearch.asp?keys=admsearch">⊙-高级搜索</a></td>
<td width="35%" align="center">
<center>现有<font color="#FF6600">    <!--#include file="zongshu.asp" -->
</font>位大侠加入</center>
</td>
</tr></table></center></div>

</td></tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>姓名：</td>
<td width="479" bgcolor="#DCE2FF">
<input name=action type=hidden value=add>
<input type="text" name="name" value="<%=uname%>" size="15" maxlength="12">
</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>门派：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="mode" value="<%=rs("门派")%>" size="15" maxlength="50">你在剑侠的门派</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>密码：</td>
<td width="479" bgcolor="#DCE2FF"><input type="password" name="password" size="15" maxlength="30">防止黑客，请不要和剑侠的登陆密码相同</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>确认密码：</td>
<td width="479" bgcolor="#DCE2FF"><input type="password" name="repassword" size="15" maxlength="30">同上</td>
</tr>

<tr>
<td width="111" align="right">性别：</td>
<td width="479" bgcolor="#DCE2FF"><select name="sex" size="1" class="p9">
<option value="男">性别→男 </option>
<option value="女">性别→女 </option>
</select></td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>信箱：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="mail" size="30" maxlength="50" value="<%=rs("电子邮箱")%>"></td>
</tr>

<tr>
<td width="111" align="right">主页：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="url" size="30" maxlength="50" value="http://"></td>
</tr>

<tr>
<td width="111" align="right">您的网络寻呼机：</td>
<td width="479" bgcolor="#DCE2FF"><select name="icq_img" class="p9" size="1">
<option value="oicq.gif" selected>OICQ</option>
<option value="icq.gif">ICQ</option>
</select>
<input type="text" name="icqnum" size="10" maxlength="10" value="<%=rs("oicq")%>">
<a href="http://www.icq.com">ICQ下载(英文)</a>，
<a href="http://www.oicq.com">OICQ下载(中文)</a></td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>您的生日：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="years" size="4" maxlength="4" value="19" class="p9">年
<input type="text" name="mons" size="2" maxlength="2">月
<input type="text" name="days" size="2" maxlength="2">日</td>
</tr>

<tr>
<td width="111" align="right">联系电话：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="likes" size="30" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">手机或呼机：</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="gsm" size="30" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">邮政编码：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="pc" size="16" maxlength="6"></td>
</tr>

<tr>
<td width="111" align="right">单位名称：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="danwei" size="50" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">详细地址：</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="address" size="50" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">照片上传：</td>
<td width="479" bgcolor="#ddE2FF"><font color=aa0000>请先加入后，再到[<a href=userre.asp>修改资料</a>]处进行操作<vinput type="file" name="photo" class="p9"> </font> </td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>近况简介：</td>
<td width="479" bgcolor="#DCE2FF"><textarea name="doc" rows="3" cols="39"><%=rs("签名档")%></textarea> </td>
</tr>

<tr>
<td width="594" colspan="2">
<center><input type="submit" value="加入" name="B1">
&nbsp;&nbsp;&nbsp;
<input type="reset" value="清除" name="B2"></center></td>
</tr>


</table></center></form></div>

</body>
</html>
<%Else
Dim name,repassword,password,sex,mail,url,icq_img,icqnum,age,gsm,pc,danwei
Dim years,mons,days,likes,address,photo,doc,mode
Dim ip,counter,n,y,r,s,f,m,sj,times,conn,sql,rs,DBPath,strSQL
Dim lists,xingzuo

name     = Server.HtmlEncode(Request.Form("name"))      '姓名
password = Server.HtmlEncode(Request.Form("password"))  '密码
repassword = Server.HtmlEncode(Request.Form("repassword"))  '确认密码
sex      = Server.HtmlEncode(Request.Form("sex"))       '性别
mail     = Server.HtmlEncode(Request.Form("mail"))      'E-Mail
URL      = Server.HtmlEncode(Request.Form("URL"))       '网站链接
icq_img  = Server.HtmlEncode(Request.Form("icq_img"))   '传呼机图片
icqnum   = Server.HtmlEncode(Request.Form("icqnum"))    '传呼机号码
AGE      = Server.HtmlEncode(Request.Form("age"))       '年龄
years    = Server.HtmlEncode(Request.Form("years"))     '生日的年
mons     = Server.HtmlEncode(Request.Form("mons"))      '生日的月
days     = Server.HtmlEncode(Request.Form("days"))      '生日的日
likes    = Server.HtmlEncode(Request.Form("likes"))     '电话
gsm      = Server.HtmlEncode(Request.Form("gsm"))       '手机
pc  	 = Server.HtmlEncode(Request.Form("pc"))   	'邮政编码
danwei   = Server.HtmlEncode(Request.Form("danwei"))    '单位
address  = Server.HtmlEncode(Request.Form("address"))   '居住地
photo    = Server.HtmlEncode(Request.Form("photo"))     '相片
doc      = Server.HtmlEncode(Request.Form("doc"))       '留言
mode     = Server.HtmlEncode(Request.Form("mode"))      '班级
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
'山羊座 （12/22-1/19） 水瓶座 （1/20-2/18） 双鱼座 （2/19-3/20） 白羊座 （3/21-4/19）
'金牛座 （4/20-5/20） 双子座 （5/21-6/20） 巨蟹座 （6/21-7/22） 狮子座 （7/23-8/22）
'处女座 （8/23-9/22） 天枰座 （9/23-10/22） 天蝎座 （10/23-11/21） 射手座 （11/22-12/21）
if days="" or years="" or mons="" then
response.write "请输入数值!<a href = javascript:history.back()>返回重输</a>"
response.end
end if
if IsNumeric(mons)=false or IsNumeric(days)=false or IsNumeric(years)=false then
response.write "您输入的不是数字。<a href = javascript:history.back()>返回重输</a>"
response.end
end if
mons=cint(mons)
days=cint(days)
years=cint(years)

if mons=>13 or mons<=0 then _
response.write "月份错误，最多只有12个月，最少也有1月。<a href = javascript:history.back()>返回重输</a>":response.end
if days=>32 or days<=0 then _
response.write "日期错误，请认真点填写！！<a href = javascript:history.back()>返回重输</a>":response.end
if years<=1899 or years>=year(now()) then _
response.write "年份是从1900年开始计算，当然，你也不能填写2000年出生的。<a href = javascript:history.back()>返回重输</a>":response.end

if (mons=12 and days>=22) or (mons=1 and days<=19) then
xingzuo= "山羊座"
elseif (mons=1 and days>=20) or (mons=2 and days<=18) then
xingzuo= "水瓶座"
elseif (mons=2 and days>=19) or (mons=3 and days<=20) then
xingzuo= "双鱼座"
elseif (mons=3 and days>=21) or (mons=4 and days<=19) then
xingzuo= "白羊座"
elseif (mons=4 and days>=20) or (mons=5 and days<=20) then
xingzuo= "金牛座"
elseif (mons=5 and days>=21) or (mons=6 and days<=20) then
xingzuo= "双子座"
elseif (mons=6 and days>=21) or (mons=7 and days<=22) then
xingzuo= "巨蟹座"
elseif (mons=7 and days>=23) or (mons=8 and days<=22) then
xingzuo= "狮子座"
elseif (mons=8 and days>=23) or (mons=9 and days<=22) then
xingzuo= "处女座"
elseif (mons=9 and days>=23) or (mons=10 and days<=22) then
xingzuo= "天枰座"
elseif (mons=10 and days>=23) or (mons=11 and days<=21) then
xingzuo= "天蝎座"
elseif (mons=11 and days>=22) or (mons=12 and days<=21) then
xingzuo= "射手座"
end if

'response.write "当前年份:" & year(now())
'response.write "<br>平均计算您的岁数是：" & (year(now()) - years)

'###############################
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("cidu-net.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
SQL="select * from list where name='"+name+"' order by ID ASC"
rs.open sql,conn,3,3
if rs.eof and rs.bof  then
elseif name = rs("name") Then
name = ""
Password=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('用户名已存在，请重新登陆填写！');" & Chr(13)
response.write "</script>" & Chr(13)
response.write "<center><font color=red><a href = javascript:history.back()>返回重输</a></font></center>"
Response.End
End If
rs.close



'--------------------判断是否写入必填字段--------
if name="" or mail="" or password="" Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('您没有填写名称或密码或信箱，请重新登陆填写！');" & Chr(13)
response.write "window.document.location.href=javascript:history.back();"&Chr(13)
response.write "</script>" & Chr(13)
End IF

if icq_img<>"0" Then
If icqnum="" Or IsNumeric(icqnum)=False Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('OICQ或icq号码输入错误!请重新输入！！！');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if

end if

if icq_img = "0" then
If icqnum <> "" Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('请选择传呼机类型！');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if
if icqnum = "" then
icqnum = "-------"
icq_img= null
end if
End If

if doc="" then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('请输入留言！');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
end if

if likes = "" then
likes="--------"
end if

if url = "" then
url="--------"
end if

if address = "" then
address="--------"
end if

if pc = "" then
pc=0
end if

if danwei = "" then
danwei="--------"
end if

if gsm = "" then
gsm="-------"
end if

if photo = "" then
photo="-------"
end if

'--------------------正确的mail判断函数----------
if IsValidEmail(mail)=False then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('错误的email输入');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if


function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
for each name in names
if Len(name) <= 0 then
IsValidEmail = false
exit function
end if
for i = 1 to Len(name)
c = Lcase(Mid(name, i, 1))
if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
IsValidEmail = false
exit function
end if
next
if Left(name, 1) = "." or Right(name, 1) = "." then
IsValidEmail = false
exit function
end if
next
if InStr(names(1), ".") <= 0 then
IsValidEmail = false
exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
IsValidEmail = false
exit function
end if
if InStr(email, "..") > 0 then
IsValidEmail = false
end if

end function

'-------------写入数据库start-------------
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("cidu-net.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
'sql="select * from list order by ID ASC"
sql="list"
rs.open sql,conn,3,3
rs.addnew
photo = null
rs.fields("name")=name
rs.fields("password")=password
rs.fields("sex")=sex
rs.fields("mail")=mail
rs.fields("url")=url
rs.fields("icqnum")=icqnum
rs.fields("icq_img")=icq_img
rs.fields("xingzuo")=xingzuo
age=n-years
rs.fields("age")=age       '年龄需要根据生日修改，而且需要制作星座代码
rs.fields("years")=years
rs.fields("mons")=mons
rs.fields("days")=days
rs.fields("likes")=likes
rs.fields("gsm")=gsm
rs.fields("pc")=pc
rs.fields("danwei")=danwei
rs.fields("address")=address
rs.fields("photo")=photo
rs.fields("doc")=doc
rs.fields("mode")=mode
rs.fields("ip")=ip
rs.fields("counter")=counter
rs.fields("times")=times
rs.update
rs.close
'-----------------------数据添加成功,页面跳转到main.asp-----------
Response.Redirect "main.asp?action=search"
End If
%>
