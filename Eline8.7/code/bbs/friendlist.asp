<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: friendlist.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim msg
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
dim ftype,Fshow
ftype=checkstr(trim(request("type")))
if ftype="" then ftype=0
if ftype=1 then
	stats="我的黑名单"
	Fshow="黑名单"
elseif ftype=0 then
	stats="好友列表"
	Fshow="好友"
else
	errmsg=errmsg+"<br>"+"<li>参数错误！"
	founderr=true
end if
call nav()
if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "的控制面板","usermanager.asp")
%>
<!--#include file="z_controlpanel.asp"-->
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<tr>
		<td align=right height=24 class=tablebody2><%
			if ftype=0 then
				response.write "<font color="&forum_body(8)&"><b>我的好友列表</b></font>"
			else
				response.write "<a href=?type=0><b>我的好友列表</b></a>"
			end if
			response.write "&nbsp;&nbsp;&nbsp;"
			if ftype=1 then
				response.write "<font color="&forum_body(8)&"><b>我的黑名单</b></font>"
			else
				response.write "<a href=?type=1><b>我的黑名单</b></a>"
			end if
		%>&nbsp;&nbsp;&nbsp;</td>
	</tr>
</table>
<%
	select case request("action")
	case "info"
		call info()
	case "addF"
		call addF()
	case "saveF"
		call saveF()
	case "删除"
		call DelFriend()
	case "清空好友"
		call AllDelFriend()
	case "清空黑名单"
		call AllDelFriend()
	case else
		call info()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub info()
%>
<form action="friendlist.asp?type=<%=ftype%>" method=post name=inbox>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
            <tr>
                <th valign=middle width="25%">姓名</td>
                <th valign=middle width="25%">邮件</td>
                <th valign=middle width="25%">主页</td>
                <th valign=middle width="10%">OICQ</td>
                <th valign=middle width="10%">发短信</td>
                <th valign=middle width="5%">操作</td>
            </tr>
<%
	set rs=server.createobject("adodb.recordset")
	if ftype=0 then
		sql="select F.*,U.useremail,U.homepage,U.oicq from Friend F inner join [user] U on F.F_Friend=U.username where F.F_username='"&trim(membername)&"' and F_type=0 order by F.f_addtime desc"
	else
		sql="select F.*,U.useremail,U.homepage,U.oicq from Friend F inner join [user] U on F.F_Friend=U.username where F.F_username='"&trim(membername)&"' and F_type=1 order by F.f_addtime desc"
	end if
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
                <tr>
                <td class=tablebody1 align=center valign=middle colspan=6>您的<%=fshow%>列表中没有任何内容。</td>
                </tr>
<%else%>
<%do while not rs.eof%>
                <tr>
                    <td align=center valign=middle class=tablebody1><a href="dispuser.asp?name=<%=htmlencode(rs("F_friend"))%>" target=_blank><%=htmlencode(rs("F_friend"))%></a></td>
                    <td align=center valign=middle class=tablebody1><a href="mailto:<%=htmlencode(rs("useremail"))%>"><%=htmlencode(rs("useremail"))%></a></td>
                    <td align=center class=tablebody1><a href="<%=htmlencode(rs("homepage"))%>" target=_blank><%=htmlencode(rs("homepage"))%></a></td>
                    <td align=center class=tablebody1><%=htmlencode(rs("oicq"))%></td>
                    <td align=center class=tablebody1><a href="usersms.asp?action=new&touser=<%=htmlencode(rs("f_friend"))%>">发送</a></td>
                <td align=center class=tablebody1><input type=checkbox name=id value=<%=rs("f_id")%>></td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
                
        <tr>
          <td align=right valign=middle colspan=6 class=tablebody2><input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">选中所有显示记录&nbsp;<input type=button name=action onclick="location.href='friendlist.asp?action=addF&type=<%=ftype%>'" value="添加<%=fshow%>">&nbsp;<input type=submit name=action onclick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除">&nbsp;<input type=submit name=action onclick="{if(confirm('确定清除所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空<%=fshow%>"></td>
                </tr>
                </table>
</form>
<%
end sub

sub delFriend()
dim delid
delid=replace(request.form("id"),"'","")
delid=replace(delid,";","")
delid=replace(delid,"--","")
delid=replace(delid,")","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"请选择相关参数。"
founderr=true
exit sub
else
	conn.execute("delete from Friend where F_username='"&trim(membername)&"' and F_id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除选定的"&fshow&"记录。"
	call dvbbs_suc()
end if
end sub
sub AllDelFriend()
	conn.execute("delete from Friend where F_username='"&trim(membername)&"' and F_type="&Ftype)
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除了所有"&fshow&"列表。"
	call dvbbs_suc()
end sub

sub addF()
%>
<form action="Friendlist.asp?type=<%=ftype%>" method=post name=messager>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=2> <input type=hidden name="action" value="saveF"> 加入<%=fshow%>--请完整输入下列信息</th>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle width=15%></td>
            <td class=tablebody1 valign=middle height=24><b>是否用短信通知对方</b><input type=checkbox name=msgto value="yes" checked></td>
          <tr> 
            <td class=tablebody1 valign=middle width=15%><b><%=fshow%>：</b></td>
            <td class=tablebody1 valign=middle height=24> <input type=text name="touser" size=50 value="<%=request("myFriend")%>">&nbsp;使用逗号（,）分开，最多5位用户</td>
          </tr>
          <tr> 
            <td valign=middle colspan=2 align=center class=tablebody2> <input type=Submit value="保存" name=Submit>&nbsp; <input type="reset" name="Clear" value="清除"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub saveF()
dim incept
if request("touser")="" then
	errmsg=errmsg+"<br>"+"<li>您忘记填写发送对象了吧。"
	founderr=true
	exit sub
else
	incept=checkStr(request("touser"))
	incept=split(incept,",")
end if

for i=0 to ubound(incept)
set rs=server.createobject("adodb.recordset")
sql="select username,usergroupid from [user] where username='"&incept(i)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>论坛没有这个用户，操作未成功。"
	founderr=true
	exit sub
end if
if ftype=1 then
	if rs(1)=1 or rs(1)=2 or rs(1)=3 or rs(1)=8 then
		errmsg=errmsg+"<br>"+"<li>您不能把管理员，超级版主，版主，贵宾加为黑名单。"
		founderr=true
		exit sub
	end if
end if
set rs=nothing

if membername=trim(incept(i)) then
	errmsg=errmsg+"<br>"+"<li>不能把自已添加为"&fshow&"。"
	founderr=true
	exit sub
end if

sql="select F_friend from friend where F_username='"&membername&"' and  F_friend='"&incept(i)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	sql="insert into friend (F_username,F_friend,F_addtime,F_type) values ('"&membername&"','"&trim(incept(i))&"',Now(),"&ftype&")"
	conn.execute(sql)
	dim msgcontent,msgtitle
	if ftype=0 then
		msgcontent="恭喜您："& chr(10) & membername &" 把您加为好友！"
	else
		msgcontent="很不幸通知您：" & chr(10) & membername &" 把您加入黑名单！"& chr(10) & "1、您将无法查看此用户的详细信息。"& chr(10) & "2、您将无法发送短信给此用户。"& chr(10) & "3、您将无法回复此用户的帖子。"
	end if
	msgtitle="通知："&membername&"把您加为"&fshow
	if Checkstr(trim(request.Form("msgto")))="yes" then
		conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&trim(incept(i))&"','"&membername&"','"&checkstr(msgtitle)&"','"&checkSTR(msgContent)&"',Now(),0,1)")
	end if
end if
if i>4 then
	errmsg=errmsg+"<br>"+"<li>每次最多只能添加5位用户，您的名单5位以后的请重新填写。"
	founderr=true
	exit sub
	exit for
end if
next
sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，"&Fshow&"添加成功。"
call dvbbs_suc()
end sub
%>