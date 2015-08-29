<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
REM 本文件改至 dvbbs 5.0 final 的 toplist.asp
'=========================================================
' File: toplist.asp
' Version:5.0
' Date: 2002-9-7
' Script Written by satan
'=========================================================
dim orders,ordername
dim Currentpage,page_count,Pcount
dim totalrec,endpage
dim GroupName,Susername,chenbackasp,chentitle
dim select1,select2,select3,select4,select5,select6,select7,select8

if Cint(GroupSetting(1))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有浏览本派弟子资料的权限，请登录或者同管理员联系。"
	founderr=true	
end if
if Request.form("action")="查找" then
	dim sqls
	Susername=Request.form("Susername")
	if Susername="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入要查找的用户名"
		founderr=true
	else
		if Request.form("search")=1 then
			sqls=" username='"&Susername&"' "	
		else
			sqls=" username like '%"&Susername&"%' "
		end if
	end if
end if
stats="论坛门派"
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()		
else
	Currentpage=request("page")
	GroupName=request("GroupName")		
			
	if not isInteger(request("orders")) or request("orders")="" then
		orders=1
	else
		orders=request("orders")
	end if
	if Request.form("action")="查找" then orders=4
	select case orders
	case 1
		orders=1
		ordername="本派所有弟子列表"
		select1="selected"
	case 2
		orders=2
		ordername="本派最新用户注册弟子列表"
		select2="selected"
	case 3
		orders=3
		ordername="本派"&Forum_Setting(68) & "大富翁"
		select3="selected"
	case 4
		orders=4
		ordername="查找本派弟子结果"
	case 7
		orders=7
		ordername="本派发帖Top"&Forum_Setting(68)
		select7="selected"
	case 8
		orders=8
		ordername="本派版主以上级别的弟子列表"
		select8="selected"
	case else
		orders=1
		ordername="本派所有弟子列表"
		select1="selected"
	end select
	
	chenbackasp="z_menpaiadd.asp?GroupName="&htmlencode(GroupName)
	chentitle=htmlencode(GroupName)
	stats=htmlencode(GroupName)&" "&ordername	

	call nav()
	call head_var(2,0,"","")
	call main()
end if
call activeonline()
call footer()

sub main()
%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<form method=POST action=z_menpaimember.asp?Groupname=<%=Groupname%>>
	<tr>
		<td colspan=5 valign=top width=350 class=tabletitle2>&nbsp;<B><%=ordername%></B></td>
		<td colspan=6 align=right class=tabletitle2>
		<select name=orders onchange='javascript:submit()'>
			<option value=1 <%=select1%>>本派所有弟子列表</option>
			<option value=2 <%=select2%>>本派最新用户注册弟子列表</option>
			<option value=3 <%=select3%>>本派<%=Forum_Setting(68)%>大富翁</option>
			<option value=7 <%=select7%>>本派发帖Top<%=Forum_Setting(68)%>排行</option>
			<option value=8 <%=select8%>>本派版主以上级别的弟子列表</option>
		</select>
		</td>
	</tr>
	</form>
	<tr> 
		<th>用户名</th>
		<th>Email</th>
		<th>OICQ</th>
		<th>主页</th>
		<th>短消息</th>
		<th>注册时间</th>
		<th>等级状态</th>
		<th>发帖总数</th>
		<th>财产</th>
	</tr>
<%
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	set rs=server.createobject("adodb.recordset")
	select case orders
	case 1
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user]  where usergroup='"&Groupname&"' order by userid desc"
	case 2
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroup='"&Groupname&"' order by AddDate desc"
	case 3
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroup='"&Groupname&"' order by userwealth desc"
	case 4
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroup='"&Groupname&"' and "&sqls&" order by username desc"
	case 7
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroup='"&Groupname&"' order by article desc"		
	case 8
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroupid<=3 and usergroup='"&Groupname&"' order by usergroupid,article desc"
	case 9
		sql="select username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where vip=1 and usergroup='"&Groupname&"' order by usergroupid,article desc"
	case else
		sql="select top "&Forum_Setting(68)&" username,useremail,userclass,oicq,homepage,article,addDate,userwealth as wealth,userid from [user] where usergroup='"&Groupname&"' order by article desc"
	end select
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=9 class=tablebody1>还没有符合要求的弟子数据</td></tr>"
	else
	totalrec=rs.recordcount
  	if totalrec mod Forum_Setting(11)=0 then
     		Pcount= totalrec \ Forum_Setting(11)
  	else
     		Pcount= totalrec \ Forum_Setting(11)+1
  	end if
	if currentpage > Pcount then currentpage = Pcount
 	if currentpage<1 then currentpage=1
	rs.PageSize = cint(Forum_Setting(11))
	rs.AbsolutePage=currentpage
	page_count=0
	do while not rs.eof and page_count < Clng(Forum_Setting(11))
%>
<tr bgcolor=<%=Forum_body(4)%>>
<td class=tablebody1>&nbsp;<a href=dispuser.asp?id=<%=rs("userid")%> target=_blank><%=rs("username")%></a></td>
<td align=center class=tablebody2><a href=mailto:<%=rs("useremail")%>><img border=0 src=<%=Forum_info(7)&Forum_TopicPic(10)%>></a></td>
<td align=center class=tablebody1>
<%=iimg(rs("oicq"),"没有","<a href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&rs("oicq")&" target=_blank><img src="&Forum_info(7)&Forum_TopicPic(11)&" alt='查看 OICQ:"&rs("oicq")&" 的资料' border=0 ></a>")%>
</td>
<td align=center class=tablebody2>
<%=iimg(rs("homepage"),"没有","<a href="&rs("homepage")&" target=_blank><img border=0 src="&Forum_info(7)&Forum_TopicPic(14)&"></a>")%>
</td>
<td align=center class=tablebody1><a href=javascript:openScript('messanger.asp?action=new&touser=<%=htmlencode(rs("username"))%>',500,400)><img src=<%=Forum_info(7)&Forum_TopicPic(7)%> border=0></a></td>
<td align=center class=tablebody2><%=rs("addDate")%></td>
<td align=center class=tablebody1> <%=rs("userclass")%><br>
</td>
<td align=center class=tablebody2><%=rs("article")%></td>
<td align=center class=tablebody1><%=rs("wealth")%></td>
</tr>
<%
		page_count=page_count+1
		rs.movenext
	loop
	end if
%>
</table>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr><td valign=middle nowrap>
<font color=<%=Forum_body(13)%>>页次：<b><%=currentpage%></b>/<b><%=Pcount%></b>页
&nbsp;
每页<b><%=Forum_Setting(11)%></b> 总数<b><%=totalrec%></b></font></td>
<td valign=middle nowrap><font color=<%=Forum_body(13)%>><div align=right><p>分页： 
<%call DispPageNum(currentpage,PCount,"""?Groupname="&Groupname&"&page=","&orders="&orders&"""")%>
</p></div></font></td></tr></table>
<%
	rs.close
	set rs=nothing
end sub
%>