<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
dim yjsk_rst,yjsk_sql,site_info,linkLogo,linkReadme,linkrs
set linkrs=server.createobject("ADODB.Recordset")
set yjsk_rst=Server.CreateObject("ADODB.RecordSet")
yjsk_sql="select * from bbslink where isnull(url) or url=''"
yjsk_rst.Open yjsk_sql,conn,1,1
if not (yjsk_rst.EOF OR yjsk_rst.BOF) then
	linkLogo=yjsk_rst("logo")
	linkreadme=yjsk_rst("readme")
	site_info=Split(yjsk_rst("boardname"),",")
	for i=0 to ubound(site_info)
		site_info(i)=clng(site_info(i))
	next
else
	linkLogo=""
	linkreadme=""
	redim site_info(6)
	site_info(0)=0
	site_info(1)=0
	site_info(2)=0
	site_info(3)=0
	site_info(4)=0
	site_info(5)=0
	site_info(6)=0
end if
if not master and not superboardmaster then
	if membername<>"" then
		if yjsk_rst.EOF OR yjsk_rst.BOF then
			errmsg=errmsg+"<br>"+"<li>管理员还没有添加本站的信息，暂时无法接受您的信息，请耐心等待！"
			founderr=true
		end if
	else
		errmsg=errmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用自助联盟功能。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
		founderr=true
	end if
end if
call nav()
stats="自助联盟"
call head_var(2,0,"","")
if founderr=true then
	call dvbbs_error()
else
	dim body
	dim readme,Tlink
	if not master and not superboardmaster then
		call main2()
	else
		call main()
	end if
	if founderr=true then	call dvbbs_error()
end if
yjsk_rst.close
set yjsk_rst=nothing
set linkrs=nothing
call activeonline()
call footer()

sub main()
	if request("action") = "add" then 
		call addlink()
	elseif request("action")="edit" then
		call editlink()
	elseif request("action")="savenew" then
		call savenew()
		if not founderr then call dvbbs_suc()
	elseif request("action")="savedit" then
		call savedit()
		if not founderr then call dvbbs_suc()
	elseif request("action")="del" then
		call del()
		if not founderr then call dvbbs_suc()
	elseif request("action")="ordersup" then
		call ordersup()
		if not founderr then response.redirect "z_link_post.asp"
	elseif request("action")="ordersdown" then
		call ordersdown()
		if not founderr then response.redirect "z_link_post.asp"
	elseif request("action")="editsiteinfo" then
		call editsiteinfo()			
	elseif request("action")="saveditsiteinfo" then
		call saveditsiteinfo()
		if not founderr then call dvbbs_suc()
	else
		call linkinfo()
	end if
end sub

sub main2()	
	dim checkuser
	if request("action")<>"del" then
		sql="select * from [bbslink] where right(lcase(trim(boardname)),"&(len(trim(membername))+1)&")='|"&lcase(trim(membername))&"'"
	else
		dim id
		id=checkstr(request("id"))
		if isnull(id) then id=""
		if id="" then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
			founderr=true
			exit sub
		end if
		sql="select * from [bbslink] where right(lcase(trim(boardname)),"&(len(trim(membername))+1)&")='|"&lcase(trim(membername))&"' and id="&id&""
	end if
	linkrs.open sql,conn,1,3
	if request("action")<>"del" then
		if site_info(0)=1 then
			if isnull(myvip) or myvip<>1 then 
				checkuser=0
			else
				checkuser=1
			end if
		else
			checkuser=1
		end if
		if checkuser=1 then
			if mymoney>=site_info(2) and myuserep>=site_info(3) and myusercp>=site_info(4) and mypower>=site_info(5) and myarticle>=site_info(6) then
				checkuser=1
			else
				checkuser=0
			end if
		end if
		if checkuser=0 then
			errmsg=errmsg+"<br>"+"<li>您当前的条件不够申请联盟论坛！"
			if site_info(0)=1 then
				errmsg=errmsg+"<br>"+"申请联盟论坛需要是VIP会员，并且用户金钱大于"&site_info(2)&"，经验大于"&site_info(3)&"，魅力大于"&site_info(4)&"，威望大于"&site_info(5)&"，帖子数大于"&site_info(6)
			else
				errmsg=errmsg+"<br>"+"申请联盟论坛需要用户金钱大于"&site_info(2)&"，经验大于"&site_info(3)&"，魅力大于"&site_info(4)&"，威望大于"&site_info(5)&"，帖子数大于"&site_info(6)
			end if
			founderr=true
		else
			if request("action")="savenew" then 
				call savenew()
				if not founderr then call dvbbs_suc()
			elseif request("action")="savedit" then
				call savedit()
				if not founderr then call dvbbs_suc()
			else
				if linkrs.bof and linkrs.eof then			
					call addlink()
				else
					call editlink() 	
				end if
			end if
		end if
	else
		if linkrs.bof and linkrs.eof then
			errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
			founderr=true
		else
			call del()
			if not founderr then call dvbbs_suc()
		end if
	end if
end sub

sub addlink()%>
	<form action="?action=savenew" method=post>
	<table cellspacing="1" cellpadding="3" class="tableBorder1" align=center>
		<tr> 
			<th width="100%" colspan=2 height=25>添加联盟论坛</th>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>论坛名称&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
	    <td width="60%" height=25 class=tablebody1 align=left>&nbsp;<input type="text" name="name" size=40></td>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>连接URL&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
	    <td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="url" size=40></td>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>连接LOGO地址&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="logo" size=40></td>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>论坛简介&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="readme" size=40></td>
	  </tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>连接方式&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="radio" name="islogo" value=0 checked>&nbsp;文字连接&nbsp;&nbsp;<%
				if master or superboardmaster or site_info(1)=1 then
					%><input type="radio" name="islogo" value=1>&nbsp;LOGO连接<%
				end if
			%></td>
		</tr>
		<tr> 
			<td height="25" colspan="2" class=tablebody2 align=center><input type="submit" name="Submit" value="添 加"></td>
		</tr>
	</table>
	</form>
<%end sub

sub editlink()
	dim boardnamesplit
	if master or superboardmaster then
		sql="select top 1 * from bbslink where id="&checkstr(Request("id"))&" and not isnull(url) and url<>''"
		linkrs.open sql,conn,1,1
	end if
	boardnamesplit=split(linkrs("boardname"),"|")
	if ubound(boardnamesplit)<1 then
		redim preserve boardnamesplit(1)
		boardnamesplit(1)=membername
	end if%>
	<form action="?action=savedit" method=post>
	<%if master or superboardmaster then%>
		<input type=hidden name=id value=<%=checkstr(Request("id"))%>>
	<%end if%>
	<table cellspacing="1" cellpadding="3" class="tableBorder1" align=center>
		<tr> 
			<th width="100%" colspan=2 height=25>编辑联盟论坛</th>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>论坛名称&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
	    <td width="60%" height=25 class=tablebody1 align=left>&nbsp;<input type="text" name="name" size=40 value=<%=htmlencode(boardnamesplit(0))%>></td>
		</tr>
		<%if master or superboardmaster then%>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>用户名称&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="username" size=40 value=<%=htmlencode(boardnamesplit(1))%>></td>
	  </tr>
	  <%end if%>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>连接URL&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
	    <td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="url" size=40 value=<%=htmlencode(linkrs("url"))%>></td>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>连接LOGO地址&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="logo" size=40 value="<%=htmlencode(linkrs("logo"))%>"></td>
		</tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>论坛简介&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="text" name="readme" size=40 value=<%=htmlencode(linkrs("readme"))%>></td>
	  </tr>
		<tr> 
			<td width="40%" height=25 class=tablebody1 align=right>连接方式&nbsp;</td>
			<td width="60%" class=tablebody1 align=left>&nbsp;<input type="radio" name="islogo" value=0 <%if linkrs("islogo")=0 then%>checked<%end if%>>&nbsp;文字连接&nbsp;&nbsp;<%
				if master or superboardmaster or site_info(1)=1 then
					%><input type="radio" name="islogo" value=1 <%if linkrs("islogo")=1 then%>checked<%end if%>>&nbsp;LOGO连接<%
				end if
			%></td>
		</tr>
		<tr> 
			<td height="25" colspan="2" class=tablebody2 align=center><input type="submit" name="Submit" value="修 改"><%if not master and not superboardmaster then%>&nbsp;&nbsp;<a href="?action=del&id=<%=linkrs("id")%>" onclick="javascript:{if(confirm('您确定执行删除操作吗?')){return true;}return false;}">[删除联盟信息]</a><input type=hidden name=username value=<%=membername%>><%end if%></td>
		</tr>
	</table>
	</form>
	<%if master or superboardmaster then
		linkrs.close
	end if
end sub

sub editsiteinfo()%>
	<form action="?action=saveditsiteinfo" method = post>
	<table cellpadding="3" cellspacing="1" class=tableborder1 align="center">
		<tr height="25"> 
			<th colspan="2">站点信息</th>
		</tr>
		<tr height="25"> 
			<td width="40%" class=tablebody1 align=right>论坛名称&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<%=forum_info(0)%></td>
		</tr>
		<tr height="25"> 
			<td width="40%" class=tablebody1 align=right>连接URL&nbsp;</td>
			<td class=tablebody1 align=left>&nbsp;<%=forum_info(1)%></td>
		</tr>
		<tr height="25"> 
			<td width="40%" class=tablebody1 align=right>连接LOGO&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input name="logo" type="text" id="logo" size=40 value="<%=linklogo%>"> [LOGO尺寸：88*31]</td>
		</tr>
		<tr height="25"> 
			<td width="40%" class=tablebody1 align=right>论坛简介&nbsp;<font color=<%=forum_body(8)%><b>*</b></font>&nbsp;</td>
			<td class=tablebody1 align=left>&nbsp;<input name="readme" type="text" id="readme" size=40 value="<%=linkreadme%>"></td>
	  </tr>
		<tr> 
			<td width="40%" class=tablebody1 align=right>联系邮箱&nbsp;</td>
			<td class=tablebody1 align=left>&nbsp;<%=forum_info(5)%></td>
		</tr>
		<tr height="25"> 
	    <td colspan=2 class=tablebody2 align=center><b>联盟要求</b></td>
		</tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>是否只允许VIP用户加入联盟&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input type="radio" name="site_info(0)" value=0 <%if site_info(0)=0 then%>checked<%end if%>>&nbsp;否&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="site_info(0)" value=1 <%if site_info(0)=1 then%>checked<%end if%>>&nbsp;是</td>
	  </tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>是否允许LOGO连接&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input type="radio" name="site_info(1)" value=0 <%if site_info(1)=0 then%>checked<%end if%>>&nbsp;否&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="site_info(1)" value=1 <%if site_info(1)=1 then%>checked<%end if%>>&nbsp;是</td>
	  </tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>用户金钱&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input name="site_info(2)" type="text" size="40" value="<%=site_info(2)%>">
	  </tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>用户经验&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input name="site_info(3)" type="text" size="40" value="<%=site_info(3)%>">
	  </tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>用户魅力&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input name="site_info(4)" type="text" size="40" value="<%=site_info(4)%>">
	  </tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>用户威望&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input name="site_info(5)" type="text" size="40" value="<%=site_info(5)%>">
	  </tr>
		<tr height="25">
	    <td width="40%" class=tablebody1 align=right>用户文章数&nbsp;</td>
	    <td class=tablebody1 align=left>&nbsp;<input name="site_info(6)" type="text" size="40" value="<%=site_info(6)%>">
	  </tr>
		<tr height="25"> 
			<td colspan="2" class=tablebody2 align="center"><input type="submit" name="Submit" value="修 改"></td>
		</tr>
	</table>
	</form>
<%end sub

sub saveditsiteinfo()
	dim linkrs
	linklogo=checkstr(Request("logo"))
	if isnull(linklogo) then linklogo=""
	linkreadme=checkstr(Request("readme"))
	if isnull(linkreadme) then linkreadme=""
	if linkreadme="" then
		errmsg=errmsg+"<br>"+"<li>论坛简介不可为空！"
		founderr=true
		exit sub
	end if
	set linkrs=Server.CreateObject("ADODB.RecordSet")
	sql="select * from bbslink where isnull(url) or url=''"
	linkrs.Open sql,conn,1,3
	if linkrs.eof or linkrs.bof then linkrs.addnew
	linkrs("boardname") = Trim(Request.Form("site_info(0)"))&","&Trim(Request.Form("site_info(1)"))&","&Trim(Request.Form("site_info(2)"))&","&Trim(Request.Form("site_info(3)"))&","&Trim(Request.Form("site_info(4)"))&","&Trim(Request.Form("site_info(5)"))&","&Trim(Request.Form("site_info(6)"))
	linkrs("readme") = Trim(linkreadme)
	linkrs("logo") = trim(linklogo)
	linkrs.update
	linkrs.Close
	set linkrs=nothing
	sucmsg="<li>您成功地修改了本站的连接信息！"
end sub

sub linkinfo()
	dim boardnamesplit
	dim totalrec,pagecount,curcount,curpage
	sql = " select * from bbslink where not isnull(url) and url<>'' order by id"
	linkrs.open sql,conn,1,1
	totalrec=linkrs.recordcount
	if totalrec mod 20=0 then
		pagecount=totalrec \ 20
	else
		pagecount=totalrec \ 20+1
	end if
	curcount=0%> 
	<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
		<tr align=center>
			<th width="8%" height=25>序号</th>
			<th width="12%" nowrap>用户名称</th>
			<th width="*">论坛名称</th>
			<th width="20%" nowrap>连接URL</th>
			<th width="5%" nowrap>连接方式</th>
			<th width="20%" nowrap>操作</th>
		</tr>
		<%curPage=request("page")
		if curpage="" or not isInteger(curpage) then
			curpage=1
		else
			curpage=clng(curpage)
		end if
		if not linkrs.eof then
			if curpage > pagecount then curpage = pagecount
			if curpage < 1 then curpage=1
			linkrs.PageSize=20
			linkrs.AbsolutePage=CurPage
			do while not linkrs.eof and curcount<20
				curcount=curcount+1
				boardnamesplit=split(linkrs("boardname"),"|")
				if ubound(boardnamesplit)<1 then
					redim preserve boardnamesplit(1)
					boardnamesplit(1)=membername
				end if
				%><tr align=center>
					<td height=25 class=tablebody1><%=linkrs("id")%></td>
					<td class=tablebody1 nowrap><%=htmlencode(boardnamesplit(1))%></td>
					<td class=tablebody1><a href=<%=linkrs("url")%> target=_blank><%=htmlencode(boardnamesplit(0))%></a></td>
					<td class=tablebody1 nowrap><a href=<%=linkrs("url")%> target=_blank><%=htmlencode(linkrs("url"))%></a></td>
					<td class=tablebody1 nowrap><%if linkrs("islogo")=1 then%>LOGO连接<%else%>文字连接<%end if%></td>
					<td class=tablebody1 nowrap><a href="?action=ordersup&id=<%=linkrs("id")%>">向上</a>&nbsp;&nbsp;<a href="?action=ordersdown&id=<%=linkrs("id")%>">向下</a>&nbsp;&nbsp;<a href="?action=edit&id=<%=linkrs("id")%>">编辑</a>&nbsp;&nbsp;<a href="?action=del&id=<%=linkrs("id")%>" onclick="javascript:{if(confirm('您确定执行删除操作吗?')){return true;}return false;}">删除</a></td>
				</tr><%
				linkrs.movenext
			loop%>
			<tr>
				<td height=25 class=tablebody2 colspan=3>&nbsp;&nbsp;页次：<b><%=curPage%></b> / <b><%=pagecount%></b> 页 每页：<b>20</b> 共计：<b><%=totalrec%></b></td>
				<td class=tablebody2 colspan=3 align=right>分页：<%
					call DispPageNum(curpage,pagecount,"?page=","")
				%>&nbsp;&nbsp;</td>
			</tr>
			<tr> 
				<td height=25 class=tablebody1 colspan=6 align=center><a href="?action=add">增加新的联盟论坛</a><%if master or superboardmaster then%> | <a href=?action=editsiteinfo><font color=<%=forum_body(8)%>>本站联盟信息</font></a><%end if%></td>
			</tr>
		<%else%>
			<tr> 
				<td height=25 class=tablebody2 colspan=6 align=center><a href="?action=add">增加新的联盟论坛</a><%if master or superboardmaster then%> | <a href=?action=editsiteinfo><font color=<%=forum_body(8)%>>本站联盟信息</font></a><%end if%></td>
			</tr>
		<%end if%>
	</table>
	<%linkrs.Close
end sub

sub savenew()
	dim mylinkname,mylinkurl,mylinkreadme
	dim islogo,mylinklogo
	mylinkname=checkstr(Request.form("name"))
	if isnull(mylinkname) then mylinkname=""
	if mylinkname="" then
		errmsg=errmsg+"<br>"+"<li>论坛名称不可为空！"
		founderr=true
	end if
	mylinkurl=checkstr(Request.form("url"))
	if isnull(mylinkurl) then mylinkurl=""
	if mylinkurl="" then
		errmsg=errmsg+"<br>"+"<li>连接URL不可为空！"
		founderr=true
		exit sub
	end if
	mylinkreadme=checkstr(Request.form("readme"))
	if isnull(mylinkreadme) then mylinkreadme=""
	if mylinkreadme="" then
		errmsg=errmsg+"<br>"+"<li>论坛简介不可为空！"
		founderr=true
		exit sub
	end if
	islogo=request.Form("islogo")
	if not isnumeric(islogo) then 
		islogo=0
	elseif islogo<>1 then
		islogo=0
	end if
	mylinklogo=checkstr(Request.form("logo"))
	if isnull(mylinklogo) then mylinklogo=""
	if islogo=1 then
		if mylinklogo="" then
			errmsg=errmsg+"<br>"+"<li>选择LOGO连接时必须指定LOGO连接的地址！"
			founderr=true
			exit sub
		end if
	end if
	if founderr then exit sub
	dim linknum,linkrs
	set linkrs= server.createobject ("adodb.recordset")
	sql = "select * from bbslink order by id desc"
	linkrs.Open sql,conn,1,3
	if linkrs.eof and linkrs.bof then
		linknum=1
	else
		linknum=linkrs("id")+1
	end if
	linkrs.Close
	set linkrs=nothing
	sql="insert into bbslink (id,boardname,readme,logo,url,islogo) values("&linknum&",'"&trim(mylinkname)&"|"&membername&"','"&trim(mylinkreadme)&"','"&trim(mylinklogo)&"','"&trim(mylinkurl)&"',"&islogo&")"
	conn.execute(sql)
	call cache_link(forum_setting(77))
	sucmsg="<li>您成功地添加了自己的联盟信息！"
end sub

sub savedit()
	dim mylinkname,mylinkurl,mylinkreadme
	dim islogo,mylinklogo,myusername
	mylinkname=checkstr(Request.form("name"))
	if isnull(mylinkname) then mylinkname=""
	if mylinkname="" then
		errmsg=errmsg+"<br>"+"<li>论坛名称不可为空！"
		founderr=true
	end if
	mylinkurl=checkstr(Request.form("url"))
	if isnull(mylinkurl) then mylinkurl=""
	if mylinkurl="" then
		errmsg=errmsg+"<br>"+"<li>连接URL不可为空！"
		founderr=true
		exit sub
	end if
	mylinkreadme=checkstr(Request.form("readme"))
	if isnull(mylinkreadme) then mylinkreadme=""
	if mylinkreadme="" then
		errmsg=errmsg+"<br>"+"<li>论坛简介不可为空！"
		founderr=true
		exit sub
	end if
	islogo=request.Form("islogo")
	if not isnumeric(islogo) then 
		islogo=0
	elseif islogo<>1 then
		islogo=0
	end if
	mylinklogo=checkstr(Request.form("logo"))
	if isnull(mylinklogo) then mylinklogo=""
	if islogo=1 then
		if mylinklogo="" then
			errmsg=errmsg+"<br>"+"<li>选择LOGO连接时必须指定LOGO连接的地址！"
			founderr=true
			exit sub
		end if
	end if
	myusername=checkstr(Request.form("username"))
	if isnull(myusername) then myusername=""
	if myusername="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
	end if
	if founderr then exit sub
	if master or superboardmaster then 
		sql = "select * from bbslink where id="&request("id")
		linkrs.Open sql,conn,1,3
	end if
	if linkrs.eof and linkrs.bof then
		errmsg=errmsg+"<br>"+"<li>错误，没有找到联盟论坛。"
		exit sub
	end if
	linkrs("boardname") = Trim(mylinkname)&"|"&trim(myusername)
	linkrs("readme") =  Trim(mylinkreadme)
	linkrs("logo")=trim(mylinklogo)
	linkrs("url") = trim(mylinkurl)
	linkrs("islogo")=islogo
	linkrs.Update
	linkrs.Close
	call cache_link(forum_setting(77))
	sucmsg="<li>您成功地修改了自己的联盟信息！"
end sub

sub del
	dim id
	id = checkstr(request("id"))
	if isnull(id) then id=""
	if id="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
	else
		sql="delete from bbslink where id="&id
		conn.Execute(sql)
		sucmsg="<li>您成功地删除了联盟信息。"
		call cache_link(forum_setting(77))
	end if
end sub

sub ordersup()
	dim id, oldid,newid
	id = checkstr(request("id"))
	if isnull(id) then id=""
	if id="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
	else
		sql="select top 2 id from bbslink where id<="&id&" and not isnull(url) and url<>'' order by id desc"
		linkrs.open sql,conn,1,3
		if not (linkrs.eof or linkrs.bof) then
			oldid=linkrs("id")
			linkrs.movenext
			if not (linkrs.eof or linkrs.bof) then
				newid=linkrs("id")
				linkrs("id")=oldid
				linkrs.update
				linkrs.MovePrevious
				linkrs("id")=newid
				linkrs.update
				call cache_link(forum_setting(77))
			end if
		end if
		linkrs.close
	end if
end sub

sub ordersdown()
	dim id, oldid,newid
	id = checkstr(request("id"))
	if isnull(id) then id=""
	if id="" then
		errmsg=errmsg+"<br>"+"<li>提交的数据有错误！"
		founderr=true
	else
		sql="select top 2 id from bbslink where id>="&id&" and not isnull(url) and url<>'' order by id"
		linkrs.open sql,conn,1,3
		if not (linkrs.eof or linkrs.bof) then
			oldid=linkrs("id")
			linkrs.movenext
			if not (linkrs.eof or linkrs.bof) then
				newid=linkrs("id")
				linkrs("id")=oldid
				linkrs.update
				linkrs.MovePrevious
				linkrs("id")=newid
				linkrs.update
				call cache_link(forum_setting(77))
			end if
		end if
		linkrs.close
	end if
end sub

%>