<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: z_school_admin.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body
stats="同学录管理页面"
call nav()
call head_var(0,0,boardtype,"z_school_class.asp?boardid="&boardid)
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>您没有在本社区同学录的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> 错误的同学录参数！请确认您是从有效的连接进入。"
	founderr=true
end if
if  boardmaster or  master or  superboardmaster then
founderr=false
else
founderr=true
Errmsg=Errmsg+"<br>"+"<li>只有管理员才能登录。"
end if
if founderr then
call dvbbs_error()
else
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
response.write "<TABLE cellpadding=0 cellspacing=1 class=tableborder1 align=center>"
response.write "<tr><td height=20 class=TopLighNav1 align=center><a href=show.asp?filetype=1&boardid="&boardid&">班级相册</a> | <a href=z_school_classuser.asp?boardid="&boardid&">班级成员</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>群体短信</a> | <a href=list.asp?boardid="&boardid&">班级论坛</a> |  "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>班级主页</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">我的资料</a> | <a href=announce.asp?boardid="&boardid&">我要留言</a> | <a href=z_school_inclass.asp?boardid="&boardid&">加入班级</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">退出班级</a> | <a href=z_school_admin.asp?boardid="&boardid&">班级管理</a></td></tr>"
response.write "<tr><td height=25 class=tablebody1 align=center>欢迎"&htmlencode(membername)&"进入同学录管理页面<td></tr><tr><td height=20 class=TopLighNav1 align=center><a href=z_school_admin.asp?action=1&boardid="&boardid&">基本信息管理</a> | <a href=z_school_admin.asp?action=2&boardid="&boardid&">班级成员管理</a> | <a href=z_school_admin.asp?boardid="&boardid&">班级公告发布</a></td></tr><tr><td class=tablebody1>"
set rs=server.createobject("adodb.recordset")
Select Case request("action")
Case 1     '基本信息管理
call editbminfo()
Case 2     '班级成员管理
call class1()
Case 3
call class2()
Case 4   '写发布(1)
call savenews()
Case 11   '基本信息管理(1)
call savebminfo()
Case 22    '班级成员管理(2)
call class2()
Case else    '写发布
call news()
end Select
if founderr then call dvbbs_error()
response.write "</td></tr></TABLE>"
end if
call activeonline()
call footer()
'------------班级成员管理-----------------------------------
sub class1()
const zdes=20
dim page,pga,pgb
if request("page")="" or not isnumeric(request("page")) then
	page=1
else
	page=clng(request("page"))
end if
dim boarduser,userzz
userzz="<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style='width:100%'><form action='z_school_admin.asp?action=22&boardid="&boardid&"' method=post name=FORM>"
set rs=conn.execute("select boarduser from board where boardid="&boardid)
if rs(0)<>"" then
boarduser=split(replace(rs(0)," ",""),",")
else
	response.write"该同学录还没有同学！"
	exit sub
end if
set rs=nothing
dim userinfo,a,pgc,pgd
pgd=ubound(boarduser)+1
pgc="本班人数"&pgd
a=1
userzz=userzz&"<tr height=25><th>操作</th><th>姓　　名</th><th>E-MAIL</th><th>电　　话</th><th>地　　　　　　址</th><th>城　　市</th></tr>"
if pgd > zdes then
pga=zdes*(page-1)
pgb=zdes*page-1
pgc=pgc&"，分页："
for i = 1 to (pgd+zdes-1)/zdes 
if i = page then
pgc=pgc&" [<font color=#FF0000>"&i&"</font>] "
else
pgc=pgc&" [<a href=z_school_admin.asp?action=2&boardid="&boardid&"&page="&i&">"&i&"</a>] "
end if
next
else
pga=0
pgb=pgd-1
end if
if pgb > ubound(boarduser) then
pgb=ubound(boarduser)
end if
dim ss,dd,ff,gg
for i = pga to pgb
if boarduser(i)<>"" then
set rs=conn.execute("select userid,useremail,userinfo from [User] where username='"&boarduser(i)&"'")
if rs(2)<>"" then
userinfo=split(rs(2),"|||")
ss=userinfo(0)
dd=userinfo(5)
ff=userinfo(13)
gg=userinfo(14)
else
ss=" "
dd=" "
ff=" "
gg=" "
end if
userzz=userzz&"<tr height=25><td class=tablebody"&a&" align=center><input type=checkbox name=boarduser value='"&boarduser(i)&"' ></td><td class=tablebody"&a&"><a href=DISPUSER.ASP?id="&rs(0)&" target=_blank>"&ss&"("&boarduser(i)&")</a></td><td class=tablebody"&a&">"&rs(1)&"</td><td class=tablebody"&a&">"&ff&"</td><td class=tablebody"&a&">"&gg&"</td><td class=tablebody"&a&">"&dd&"</td></tr>"
set rs=nothing
end if
if a=1 then a=2 else a=1
next
userzz=userzz&"<tr height=25><td class=tablebody1 colspan=5>"&pgc&"</td><td class=tablebody2 align=center><input type=Submit value='开 除' name=Submit></td></tr></form></TABLE>"
response.write userzz
end sub
'------------班级成员管理(2)-----------------------------------
sub class2()
dim boarduser,username
username=replace(request("boarduser")," ","")
sql="select boarduser,txlpd from board where boardid="+Cstr(request("boardid"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>错误版面参数！"
	founderr=true
	exit sub
elseif cint(left(rs("txlpd"),1))=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误同学录版面参数！"
	founderr=true
	exit sub
elseif isnull(rs(0)) or rs(0)="" then
	Errmsg=Errmsg+"<br>"+"<li>错误同学录参数！"
	founderr=true
	exit sub
else
	boarduser=rs(0)
end if
if instr(username,",") = 0 then
	username=" "&username&" "
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	boarduser=replace(boarduser,username," ")
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
else
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	dim dede
	username=split(username, ",")
	for each dede in username
		dede=" "&dede&" "
		boarduser=replace(boarduser,dede," ")
	next
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
end if
conn.execute("update board set boarduser='"&boarduser&"' where boardid="+Cstr(request("boardid")))
response.write "开除成功！"&request("boarduser")
end sub
'-----------------------------------------------
sub class3()



end sub
'------------------写发布-----------------------------
sub news()
%>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:100%">
<form action="z_school_admin.asp?action=4" method=post name=FORM>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">发布人： </div>
      </td>
      <td width="80%" class=tablebody1>
        <input type=text name=username size=36 value="<%=membername%>" disabled>
        <input type=hidden name=username value="<%=membername%>">
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">标题： </div>
      </td>
      <td width="80%" class=tablebody1>
        <input type=text name=title size=60><input type=hidden name="boardid" value="<%=CStr(BoardID)%>">
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">内容： </div>
      </td>
      <td width="80%" class=tablebody1>
        <textarea cols=60 rows=6 name="content"></textarea>
      </td>
    </tr>
    <tr>
      <td width="100%" valign=top colspan="2" align=center class=tablebody2> 
        <input type=Submit value="发 送" name=Submit>
        &nbsp; 
        <input type="reset" name="Clear" value="清 除">
      </td>
    </tr>
  </table>
<%
end sub
'-------------------写发布(1)----------------------------
sub savenews()
	dim username,title,content
	if request("boardid")<>"" or (not isInteger(request("boardid"))) then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br>"+"<li>您输入了错误的参数。"
		founderr=true
	end if
	if request("username")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入发布者。"
		founderr=true
	else
		username=checkstr(request("username"))
	end if
	if request("title")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入新闻标题。"
		founderr=true
	else
		title=checkstr(request("title"))
	end if
	if request("content")="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入新闻内容。"
		founderr=true
	else
		content=checkstr(request("content"))
	end if

	if founderr=true then
		call dvbbs_error()
		exit sub
	end if 
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(25))=1 then
		sql="select * from bbsnews"
		rs.open sql,conn,1,3
		rs.addnew
		rs("username")=username
		rs("title")=title
		rs("content")=content
		rs("addtime")=Now()
		rs("boardid")=boardid
		rs.update
		rs.close
		myCache.name="AnnounceMents"&BoardID
		myCache.makeEmpty
		call success()
		else
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
	end if
end sub
sub success()
%><br><br>
成功：新闻操作
<%
end sub

'基本信息管理
sub editbminfo()
dim master_1

%> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="+CSTr(boardid)
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
   call dvbbs_error()
exit sub
   end if
master_1=split(rs("boardmaster"),"|")
if not master then
	if membername<>master_1(0)  then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主管理员专用。"
	call dvbbs_error()
	exit sub
	else
	Errmsg=Errmsg+"<br>"+"<li>您未有修改设置的权限。"
	call dvbbs_error()
	exit sub
	end if
end if
Forum_Setting = split(rs("Board_Setting"),",")
%>
             
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:100%">
<form action ="z_school_admin.asp?action=11&boardid=<%=boardid%>" method=post>
    <tr> 
      <th width="20%" height=22 class=tablebody2><b>字段名称：</b> </td>
      <th  > 
        <div align="center"><b>变量值：</b></div>
      </th>
    </th>
    <tr> 
      <td height=22 class=tablebody1  align="center">班级名称：</td>
      <td  class=tablebody1>
	  <input type="text" name="boardtype" size="30" value='<%=htmlencode(rs("boardtype"))%>'>
	  <input type='hidden' name=editid value='<%=boardid%>'></td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2  align="center">版面说明：</td>
      <td  class=tablebody1> 
        <input type="text" name="readme" size="60" value='<%=htmlencode(rs("readme"))%>'>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody1  align="center">副管理员：</td>
      <td  class=tablebody1> 
        <input type="text" name="boardmaster" size="30" value='<%
if instr(rs("boardmaster"),"|") <> 0 then
response.write replace(rs("boardmaster"),master_1(0)&"|","")
end if
%>'>(多副管理员添加请用|分隔，如：追梦人|wxz)
      </td>
    </tr>
<tr> 
      <td height=22 class=tablebody1  align="center">是否开放：</td>
      <td  class=tablebody1>
<input type=radio name="Board_Setting(2)" value="0" <%if cint(Forum_Setting(2))=0 then%>checked<%end if%>>是&nbsp;
<input type=radio name="Board_Setting(2)" value="1" <%if cint(Forum_Setting(2))=1 then%>checked<%end if%>>否&nbsp;(开放则不是本班的同学也可以进来!)
      </td>
    </tr><%
dim txlpd
txlpd=split(rs("txlpd"),"$")
if txlpd(1)="" then
txlpd(1)=0
end if
%>
<tr> 
      <td height=22 class=tablebody1  align="center">入学年份：</td>
      <td  class=tablebody1>
<select name="txlpd(1)">
<%for i = 1955 to 2003%>
<option value="<%=i%>" <%if i=cint(txlpd(1)) then%>selected<%end if%>><%=i%></option>
<%next%>
</select>
      </td>
    </tr>
<tr> 
      <td height=22 class=tablebody1  align="center">班级主页：</td>
      <td  class=tablebody1>http://<input type="text" name="txlpd(2)" size="30" value='<%=txlpd(2)%>'>
     </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2>&nbsp;</td>
      <td  class=tablebody2> 
        <input type="submit" name="Submit" value="提交">
      </td>
    </tr></form>
  </table>
<%
rs.close
end sub


'基本信息管理(1)
sub savebminfo()
dim rname,zmzm,boarduser,zmrr,txlpd,txlpd1
set rs= server.CreateObject ("adodb.recordset")
sql = "select boardmaster,Board_Setting,boarduser,txlpd from board where boardid="+CSTr(request("boardid"))
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
   call dvbbs_error()
exit sub
end if
if not master then
	rname=split(rs(0),"|")
	if membername<>rname(0)  then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主管理员专用。"
	call dvbbs_error()
	exit sub
	else
	Errmsg=Errmsg+"<br>"+"<li>您未有修改设置的权限。"
	call dvbbs_error()
	exit sub
	end if
end if
txlpd=split(rs(3),"$")
zmrr=split(checkStr(rs(0)),"|")
boarduser=rs(2)
zmzm=rs(1)
set rs=nothing
if CSTr(Request("Board_Setting(2)"))=1 then
zmzm=left(zmzm,4)&"1"&mid(zmzm,6)
elseif CSTr(Request("Board_Setting(2)"))=0 then
zmzm=left(zmzm,4)&"0"&mid(zmzm,6)
end if
rname=split(checkStr(Request("boardmaster")),"|")
for i=0 to ubound(rname)
sql="select top 1 username from [user] where username='"&replace(rname(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>论坛没有这个用户，不能添加为副管理员"
	exit sub
	exit for
elseif instr(boarduser,","&rname(i))=0 and instr(boarduser,rname(i)&",")=0 then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>同学录没有这个用户，不能添加为副管理员"
	exit sub
	exit for
end if
set rs=nothing
next
txlpd1=txlpd(0)&"$"&checkreal(request.Form("txlpd(1)"))&"$"&checkreal(request.Form("txlpd(2)"))&"$"&txlpd(3)&"$"&txlpd(4)&"$"&txlpd(5)&"$1"
rname = zmrr(0)&"|"&checkStr(Request("boardmaster"))
set rs=server.createobject("adodb.recordset")
sql = "select * from board where boardid="+Cstr(request("boardid"))
rs.open sql,conn,1,3
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
if Request("boardmaster")<>"" then
	rs("boardmaster") = rname
end if
	rs("readme") = checkStr(Request("readme"))
	rs("boardtype")=checkStr(Request("boardtype"))
	rs("Board_Setting")= zmzm
	rs("txlpd")= txlpd1
	rs.Update 
	rs.Close 
	response.write "<p>论坛修改成功！"
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"$","§")
	checkreal=w
end if
end function
%>




