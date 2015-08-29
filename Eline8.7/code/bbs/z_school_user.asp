<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/z_school_char.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/birthday.asp"-->
<%
'=========================================================
' File: z_school_user.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'========================================================= 
dim suserid
if request("id")="" then
	suserid=userid
elseif isnumeric(request("id")) then
	suserid=request("id")
else
	suserid=session("userid")
end if
stats="成员资料"
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)

if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>您没有在本社区同学录的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> 错误的同学录参数！请确认您是从有效的连接进入。"
	founderr=true
elseif cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本同学录为认证同学录，请<a href=login.asp>登录</a>并确认您是本班同学。"
		founderr=true
	else
		if chkschoollogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>你不是本班同学，请先<a href=z_school_inclass.asp?boardid="&boardid&">加入本班</a>。"
		founderr=true
		end if
	end if
end if
if founderr then
call dvbbs_error()
else
call class1()
end if
call activeonline()
call footer()
sub class1()
dim boarduser,userzz
response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
response.write "<col width=20% ><col width=*><col width=40% >"
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
response.write "<tr><td height=20 class=TopLighNav1 align=center colspan=3><a href=show.asp?filetype=1&boardid="&boardid&">班级相册</a> | <a href=z_school_classuser.asp?boardid="&boardid&">班级成员</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>群体短信</a> | <a href=list.asp?boardid="&boardid&">班级论坛</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>班级主页</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">我的资料</a> | <a href=announce.asp?boardid="&boardid&">我要留言</a> | <a href=z_school_inclass.asp?boardid="&boardid&">加入班级</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">退出班级</a> | <a href=z_school_admin.asp?boardid="&boardid&">班级管理</a></td></tr>"
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where userid="&suserid
rs.open sql,conn,1,1
if err.number<>0 then 
	ErrMsg=Errmsg+"<br>"+"<li>数据库操作失败："&err.description
	founderr=true
	exit sub
end if
if rs.eof and rs.bof then
	ErrMsg=Errmsg+"<br>"+"<li>您查询的名字不存在"
	founderr=true
	exit sub
end if
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,useraddress,userinfo
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		useraddress=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		useraddress=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	useraddress=""
end if
%> 
  <tr height=25> 
    <th colspan=2 align=left>基本资料</th>
    <td rowspan=9 align=center class=tablebody1 width=40% valign=top>
<%=iimg(rs("userphoto"),"<font color=gray>无</font>","<img src="&htmlencode(FilterJS(rs("userPhoto")))&">")%>
    </td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>名 称：</td>
    <td class=tablebody2><%=htmlencode(rs("username"))%></td>
</tr>   
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>性 别：</td>
    <td class=tablebody1><%=iif(rs("sex"),"男","女")%> </td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>出 生：</td>
    <td class=tablebody2>
<%=iimg(rs("birthday"),"<font color=gray>未填</font>",rs("birthday"))%>
 </td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>星 座：</td>
    <td class=tablebody1>
<%=iimg(rs("birthday"),"<font color=gray>未填</font>",astro(rs("birthday")))%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>Ｅｍａｉｌ：</td>
    <td class=tablebody2>
<%=iimg(rs("useremail"),"<font color=gray>未填</font>","<a href=mailto:"&htmlencode(rs("useremail"))&">"&htmlencode(rs("useremail"))&"</a>")%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>Ｑ Ｑ：</td>
    <td class=tablebody1>
<%=iimg(rs("oicq"),"<font color=gray>未填</font>",htmlencode(rs("oicq")))%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>ＩＣＱ：</td>
    <td class=tablebody2>
<%=iimg(rs("icq"),"<font color=gray>未填</font>",htmlencode(rs("icq")))%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>ＭＳＮ：</td>
    <td class=tablebody1>
<%=iimg(rs("msn"),"<font color=gray>未填</font>",htmlencode(rs("msn")))%>
 </td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>主 页：</td>
    <td class=tablebody2>
<%=iimg(rs("homepage"),"<font color=gray>未填</font>","<a href="&htmlencode(rs("homepage"))&" target=_blank>"&htmlencode(rs("homepage"))&"</a> ")%>
</td><td class=tablebody1 align=center width=40% >
      <b><a href="javascript:openScript('messanger.asp?action=new&touser=<%=htmlencode(rs("username"))%>',500,400)">给他留言</a> | <a href="friendlist.asp?action=addF&myFriend=<%=HTMLEncode(rs("username"))%>" target=_blank>加为好友</a></b></td>
  </tr>
<tr height=25> 
    <th colspan=2 align=left>
      用户详细资料</th>
    <td rowspan=14 class=tablebody1 width=40% valign=top>
<b>性格：</b>
<br>
<%=character%>
<br><br><br>
<b>个人简介：</b><br>
<%=personal%>
<br>
</td>
  </tr>   
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>真实姓名：</td>
    <td class=tablebody1><%=realname%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>国　　家：</td>
    <td class=tablebody2><%=country%> </td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>联系电话：</td>
    <td class=tablebody1><%=userphone%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>通信地址：</td>
    <td class=tablebody2><%=useraddress%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>省　　份：</td>
    <td class=tablebody1><%=province%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>城　　市：</td>
    <td class=tablebody2><%=city%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>生　　肖：</td>
    <td class=tablebody1><%=shengxiao%> </td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>血　　型：</td>
    <td class=tablebody2><%=blood%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>信　　仰：</td>
    <td class=tablebody1><%=belief%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>职　　业：</td>
    <td class=tablebody2><%=occupation%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>婚姻状况：</td>
    <td class=tablebody1><%=marital%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>最高学历：</td>
    <td class=tablebody2><%=education%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>毕业院校：</td>
    <td class=tablebody1><%=college%></td>
  </tr>
<%
rs.close
set rs=nothing
call activeonline()
response.write "</TABLE>"
end sub
%>