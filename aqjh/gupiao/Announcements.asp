<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-股市公告</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0>
<table  width="97%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>浏览股市公告</b></font></td>
	</tr>
</table>	
<% 
	select case request("action")
		case "AddAnn"		
			call AddAnn()
		case "SaveAnn"		
			call SaveAnn()
		case "EditAnn"
			call EditAnn()
		case "EditAnnInfo"
			call EditAnnInfo()		
		case "SaveEdit"		
			call SaveEdit()
		case "DelAnn"
			call DelAnn()							
		case else
			call main()
	end select
%>
<br>
<table width="97%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
<tr><td height=32 background="images/footer.gif" valign=middle align="center"><A href="stock.asp" class="cblue">[返回股市大厅]</A></td></tr>
</table>
</body></html>
<%
CloseDatabase		'关闭数据库
'=====================================
sub main()
	if master then
		response.write "<br><table border=0 cellpadding=0 cellspacing=0 align=center><tr><td><a href=?action=AddAnn>发布公告</a> | <a href=?action=EditAnn>管理公告</a> | <a href=Admin_Setting.asp>股市管理</a></td></tr></table>"
	end if
%>
<br>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
<%
	set rs=conn.execute("select * from [StockNews] order by id desc")
	if rs.eof and rs.bof then
		response.Write "<tr><td colspan=2 height=""25"" class=""forumRow""><font color=gray>当前没有公告</font></td></tr>"
	else
		do while not rs.eof	
%>
	<tr>
		<th height=25 class=admin>>> <%=server.htmlencode(rs("title"))%> <<</th>
	</tr>	
	<tr>
		<td class="forumRow" valign=top style='LEFT: 0px; WIDTH: 100%; WORD-WRAP: break-word'>
		<br><p><blockquote><%=rs("content")%></blockquote></p>
		</td>  
	</tr>
	<tr>
		<td align=left class=bg1>
			<table width=100% border=0 cellpadding=0 cellspacing=0>
			<tr><td align=left>&nbsp;&nbsp;&nbsp;<b>发布人</b>： <%=server.htmlencode(rs("username"))%>
			</td><td align=right><b>发布时间</b>： <%=rs("addtime")%>&nbsp;&nbsp;&nbsp;
			</tr></table>
		</td>	
	</tr>
			
<%
		rs.movenext
		loop
	end if	
%>																						
</table>	
<% 
end sub
'-----------------------
sub AddAnn()
	if not master then
		Errmess=Errmess+"<br><li>您没有执行此操作的权限。"
		call endinfo(1)
		exit sub
	end if
	response.write "<br><table border=0 cellpadding=0 cellspacing=0 align=center><tr><td><a href=?>浏览公告</a> | <a href=?action=AddAnn>发布公告</a> | <a href=?action=EditAnn>管理公告</a> | <a href=Admin_Setting.asp>股市管理</a></td></tr></table>"
%>
<br>
<table cellpadding=6 cellspacing=1 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
    <tr><th valign=middle colspan=2 class="admin">发布公告</th></tr>
<form action="?action=SaveAnn" method="post" name="form2"> 	
    <tr>
    	<td class=forumRow valign=middle><b>用户名：</b></td>
		<td class=forumRow valign=middle>
				<%=membername%><INPUT name=username type=hidden value="<%=membername%>">
		</td>
	</tr>
    <tr>
   	 	<td class=forumRow valign=middle><b>标 题：</b><br>100字以内</td>
    	<td class=forumRow valign=middle><INPUT name="title" type=text size=60></td>
	</tr>
    <tr>
   	 	<td class=forumRow valign=middle><b>颜色：</b><br>公告标题颜色</td>
		<td class=forumRow valign=middle>
			<SELECT name=color> 
			<option style="background-color:black;color: black" value="black" selected> black </option>
			<option style="background-color:blue;color: blue" value="blue"> blue </option>
			<option style="background-color:red;color: red" value="red"> red </option>
			<option style="background-color:olive;color: olive" value="olive"> olive </option>
			<option style="background-color:teal;color: teal" value="teal"> teal </option>
			<option style="background-color:maroon;color: maroon" value="maroon"> maroon </option>
			<option style="background-color:navy;color: navy" value="navy"> navy</option>
			<option style="background-color:gray;color: gray" value="gray"> gray</option>
			<option style="background-color:lime;color: lime" value="lime"> lime</option>
			<option style="background-color:Fuchsia;color: Fuchsia" value="Fuchsia"> Fuchsia </option>
			<option style="background-color:yellow;color: yellow" value="yellow"> yellow </option>
			<option style="background-color:green;color: green" value="green"> green</option>
			<option style="background-color:Purple;color: Purple" value="Purple"> Purple</option>
			</SELECT>
		</td>
	</tr>
    <tr>
    	<td class=forumRow valign=top width=30%><b>内 容：</b><br>250字以内<br><br><font color="red">HTML语法 支持<br>UBB 语法 不支持</font></td>
		<td class=forumRow valign=middle>
			<textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"></textarea>
		</td>
	</tr>
    <tr>
    	<td class=forumRow valign=middle colspan=2 align=center><input type=submit name="submit" value="发 布"></td>
	</tr>
</form>
</table>
<%
end sub

sub SaveAnn()
	dim username,title,content
	if not master then
		Errmess=Errmess+"<br><li>您没有执行此操作的权限。"
		call endinfo(1)
		exit sub
	end if
	if request("username")="" then
		Errmess=Errmess+"<br>"+"<li>请输入发布者。"
		founderr=true
	else
		username=replace(request("username"),"'","")
	end if
	if request("title")="" then
		Errmess=Errmess+"<br>"+"<li>请输入新闻标题。"
		founderr=true
	else
		title=replace(request("title"),"'","")
	end if
	if request("content")="" then
		Errmess=Errmess+"<br>"+"<li>请输入新闻内容。"
		founderr=true
	else
		content=replace(request("content"),"'","")
	end if
	if founderr then 
		call endinfo(1)
		exit sub
	end if
	conn.execute("insert into [StockNews] (title,content,[username],color,addtime) values('"&title&"','"&content&"','"&username&"','"&request.Form("color")&"',now())")	
	sucmess="<li>您已经成功的发布了公告。"
	call endinfo(1)
end sub
'---------------管理公告--------------------
sub EditAnn()
	if not master then
		Errmess=Errmess+"<br><li>您没有执行此操作的权限。"
		call endinfo(1)
		exit sub
	end if
	response.write "<br><table border=0 cellpadding=0 cellspacing=0 align=center><tr><td><a href=?>浏览公告</a> | <a href=?action=AddAnn>发布公告</a> | <a href=?action=EditAnn>管理公告</a> | <a href=Admin_Setting.asp>股市管理</a></td></tr></table>"
%>
<br>
<table width="97%" cellpadding=6 cellspacing=1 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
    <tr>
     	<th class="admin">ID</th>
		<th class="admin">标题</th>
		<th class="admin">发布人</th>
		<th class="admin">发布时间</th>
		<th class="admin">操作</th>
	</tr>
<form action="?action=DelAnn" method="post" name="form1"> 	
<%
	dim hf
	set rs=conn.execute("select * from [StockNews] order by id desc")
	if rs.eof and rs.bof then
		hf=0
		response.Write "<tr><td colspan=5 height=""25"" class=""forumRow""><font color=gray>当前没有公告</font></td></tr>"
	else
		hf=1
		do while not rs.eof	
%>	
    <tr align="center">
    	<td class=forumRow valign=middle><%=rs("id")%></td>
		<td class=forumRow valign=middle align="left"><a href="?action=EditAnnInfo&id=<%=rs("id")%>"><%=server.htmlencode(rs("title"))%></a></td>
    	<td class=forumRow valign=middle><%=server.htmlencode(rs("username"))%></td>
		<td class=forumRow valign=middle><%=rs("addtime")%></td>
    	<td class=forumRow valign=middle><input type="checkbox" name="id" value="<%=rs("id")%>"></td>				
	</tr>
<%
		rs.movenext
		loop
	end if
	rs.close
	if hf=1 then	
%>	
	<tr><td class=bg1 colspan=5 align=right height="25">请选择要删除的公告，<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">全选 <input type=submit name=Submit value=执行  onclick="{if(confirm('您确定执行删除操作吗?')){this.document.form1.submit();return true;}return false;}"></td></tr>
<%  end if%>
</form>
</table>	

<%
end sub
'--------------编辑公告--------------
sub EditAnnInfo()
	if not master then
		Errmess=Errmess+"<br><li>您没有执行此操作的权限。"
		call endinfo(1)
		exit sub
	end if
	if not isnumeric(request("id")) then
		Errmess=Errmess+"<br><li>错误的参数。"
		call endinfo(1)
		exit sub
	end if	
	response.write "<br><table border=0 cellpadding=0 cellspacing=0 align=center><tr><td><a href=?>浏览公告</a> | <a href=?action=AddAnn>发布公告</a> | <a href=?action=EditAnn>管理公告</a> | <a href=Admin_Setting.asp>股市管理</a></td></tr></table>"
	set rs=conn.execute("select * from [StockNews] where id="&request("id"))
%>
<br>
<table cellpadding=6 cellspacing=1 align=center style="border: 1px #6595D6 solid; background-color: #FFFFFF;">
    <tr>
    <th valign=middle colspan=2 class="admin">管理股市公告</th></tr>
<form action="?action=SaveEdit&id=<%=request("id")%>" method="post"> 	
    <tr>
    	<td class=forumRow valign=middle><b>用户名：</b></td>
		<td class=forumRow valign=middle>
		<input type=text name=username value="<%=rs("username")%>">
		</td>
	</tr>
    <tr>
   	 	<td class=forumRow valign=middle><b>标 题：</b><br>100字以内</td>
    	<td class=forumRow valign=middle><INPUT name="title" type=text size=60 value="<%=rs("title")%>"></td>
	</tr>
    <tr>
   	 	<td class=forumRow valign=middle><b>颜色：</b><br>公告标题颜色</td>
    	<td class=forumRow valign=middle>
<SELECT name=color> 
<option style="background-color:black;color: black" value="black" <%if rs("color")="black" then%> selected <%end if%>> black </option>
<option style="background-color:blue;color: blue" value="blue" <%if rs("color")="blue" then%> selected <%end if%>> blue </option>
<option style="background-color:red;color: red" value="red" <%if rs("color")="red" then%> selected <%end if%>> red </option>
<option style="background-color:olive;color: olive" value="olive" <%if rs("color")="olive" then%> selected <%end if%>> olive </option>
<option style="background-color:teal;color: teal" value="teal" <%if rs("color")="teal" then%> selected <%end if%>> teal </option>
<option style="background-color:maroon;color: maroon" value="maroon" <%if rs("color")="maroon" then%> selected <%end if%>> maroon </option>
<option style="background-color:navy;color: navy" value="navy" <%if rs("color")="navy" then%> selected <%end if%>> navy</option>
<option style="background-color:gray;color: gray" value="gray" <%if rs("color")="gray" then%> selected <%end if%>> gray</option>
<option style="background-color:lime;color: lime" value="lime" <%if rs("color")="lime" then%> selected <%end if%>> lime</option>
<option style="background-color:Fuchsia;color: Fuchsia" value="Fuchsia" <%if rs("color")="Fuchsia" then%> selected <%end if%>> Fuchsia </option>
<option style="background-color:yellow;color: yellow" value="yellow" <%if rs("color")="yellow" then%> selected <%end if%>> yellow </option>
<option style="background-color:green;color: green" value="green" <%if rs("color")="green" then%> selected <%end if%>> green</option>
<option style="background-color:Purple;color: Purple" value="Purple" <%if rs("color")="Purple" then%> selected <%end if%>> Purple</option>
</SELECT>
		</td>
	</tr>
    <tr>
    	<td class=forumRow valign=top width=30%><b>内 容：</b><br>250字以内<br><br><font color="red">HTML语法 支持<br>UBB 语法 不支持</font></td>
		<td class=forumRow valign=middle>
			<textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"><%=replace(replace(rs("content"),"<br>",chr(13)),"&nbsp;"," ")%></textarea>
		</td>
	</tr>
    <tr>
    <td class=forumRow valign=middle colspan=2 align=center><input type=submit name="submit" value="发 布"></td></tr>
</form>
</table>
<%
end sub

sub SaveEdit()
	dim username,title,content
	if not master then
		Errmess=Errmess+"<br><li>您没有执行此操作的权限。"
		call endinfo(1)
		exit sub
	end if
	if not isnumeric(request("id")) then
		Errmess=Errmess+"<br><li>错误的参数。"
		call endinfo(1)
		exit sub
	end if	
	if request("username")="" then
		Errmess=Errmess+"<br>"+"<li>请输入发布者。"
		founderr=true
	else
		username=replace(request("username"),"'","")
	end if
	if request("title")="" then
		Errmess=Errmess+"<br>"+"<li>请输入新闻标题。"
		founderr=true
	else
		title=replace(request("title"),"'","")
	end if
	if request("content")="" then
		Errmess=Errmess+"<br>"+"<li>请输入新闻内容。"
		founderr=true
	else
		content=replace(request("content"),"'","")
	end if
	if founderr then 
		call endinfo(1)
		exit sub
	end if
	conn.execute("update [StockNews] set title='"&title&"',content='"&content&"',[username]='"&username&"',color='"&request.Form("color")&"',addtime=now() where id="&request("id"))	
	sucmess="<li>您已经成功的修改了公告。"
	rurl="?action=EditAnn"
	call endinfo(1)
end sub

sub delann()
	if not master then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
		exit sub
	end if
	dim delid
	delid=replace(request.form("id"),"'","")
	if delid<>"" then
		conn.execute("delete from [StockNews] where id in ("&delid&")")
	end if
	sucmess="<li>删除公告成功。"
	rurl="?action=EditAnn"
	call endinfo(1)
end sub
%>