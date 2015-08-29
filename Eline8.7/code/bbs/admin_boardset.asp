<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%

	stats="版主管理页面"
	dim sql1,rs1
	if not founduser  then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
	end if

	if request("boardid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定论坛版面。"
	elseif not isInteger(request("boardid"))  or request("boardid")=0  then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的版面参数。"
	else
		boardid=clng(request("boardid"))
	end if

	if not(  boardmaster or  master or  superboardmaster ) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>只有管理员才能登录。"
	end if

	call nav()
if founderr then
	call head_var(1,BoardDepth,0,0)
	call dvbbs_error()
else
	call head_var(1,BoardDepth,0,0)
	dim title
	dim content
	set rs=server.createobject("adodb.recordset")
	call main()
end if

	set rs=nothing
	call footer()

	sub main()
%>
<TABLE cellpadding=0 cellspacing=1 class=tableborder1 align=center > 
        <tr >
          <th height=24 align=center colspan="2">欢迎 <%=htmlencode(membername)%>进入版主管理页面</th>
        </tr>
        <tr >
          <td height=24 align=center colspan="2" class=tablebody1>
        <b>管理选项：<a href=admin_boardset.asp?boardid=<%=boardid%>>论坛公告发布</a>
        <%if master then%>
        | <b><a href=admin_boardset.asp?action=manage&boardid=<%=boardid%>>公告管理</a>
        <%end if%> | 
		<a href=admin_boardset.asp?action=editbminfo&boardid=<%=boardid%>>基本信息管理</a> | 
		<a href=admin_boardset.asp?action=editbmset&boardid=<%=boardid%>>基本设置管理</a> | 
		<!--<a href=admin_boardset.asp?action=editbmset&boardid=<%=boardid%>>基本设置管理</a> | -->
		<a href=admin_boardset.asp?action=editbmcolor&boardid=<%=boardid%>>颜色设置管理</a> | 
		<!--<a href=admin_boardset.asp?action=editbmads&boardid=<%=boardid%>>分版广告管理</a> | -->
		<a href=admin_boardset.asp?action=editbmads&boardid=<%=boardid%>>分版广告管理</a>
		</b></td>
        </tr>
		<tr>
              <td width="30%" valign=top align=center class=tablebody1 >
		<table cellpadding=3 cellspacing=1   class=tableborder2 style="width:90%;word-break:break-all;" >
		<tr>
			<th width="100%" height=24  colspan="2">《 本版信息栏 》
			</th>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 colspan="2" align=center ><%=boardtype%>
			</td>
		</tr>
		<tr>
			<td width="60%" height=24 class=tablebody1 >今日新帖：
			</td>
			<td width="40%" height=24 class=tablebody1 ><FONT COLOR=RED><%=todaynum%></FONT>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 >主题帖子：
			</td>
			<td  height=24 class=tablebody2 ><%=LastTopicNum%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >本版帖子：
			</td>
			<td  height=24 class=tablebody1 ><%=LastBbsNum%>
			</td>
		</tr>
		<tr>
			<td width="30%" height=24 class=tablebody2 colspan="2">管理成员：
		<%=boardmasterlist%>
			</td>
		</tr>
		<tr>
			<th width="100%" height=24  colspan="2">《 管理权限 》
			</th>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >主版主可增删副版主：
			</td>
			<td  height=24 class=tablebody1 ><%if Board_Setting(33)=1 then%>打开<%else%><FONT COLOR=RED>关闭</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 >主版主可修改颜色配置：
			</td>
			<td  height=24 class=tablebody2 ><%if Board_Setting(34)=1 then%>打开<%else%><FONT COLOR=RED>关闭</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >所有版主均可修改颜色配置：
			</td>
			<td  height=24 class=tablebody1 ><%if Board_Setting(35)=1 then%>打开<%else%><FONT COLOR=RED>关闭</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td width="100%" height=24  colspan="2" class=tablebody2>
		<b>注意：</b>各个版面版主可以在自己版面自由发布公告和版面设置，管理员可以在所有版面发布，并对信息进行管理操作。
			</td>
		</tr>
		</table>
	      </td>
              <td width="70%" valign=top class=tablebody1 align=center>
      		<table cellpadding=3 cellspacing=1  class=tableborder2 style="width:100%;word-break:break-all;">
		  <tr>
			<td width="100%" height=24 class=tablebody2 >
<B>注意</B>：<BR>本页面为版主专用，使用前请看左边相对应的功能是否打开，在进行管理设置的时候，不要随意更改设置，如需更改，必须填写完整或者正确的填写。
		  </td></tr>
		</table>
<%
	select case request("action")
	case "new"
		call savenews()
	case "manage"
		call manage()
	case "edit"
		call edit()
	case "updat"
		call update()
	case "del"
		call del()
	case "editbminfo"
		call editbminfo()
	case "saveditbm"
		call savebminfo()
	case "editbmset"
		call editbmset()
	case "savebmset"
		call savebmset()
	case "editbmcolor"
		call editbmcolor()
	case "savebmcolor"
		call savebmcolor()
	case "editbmads"
		call editbmads()
	case "savebmads"
		call savebmads()
	case else
	call news()
	end select
	if founderr then call dvbbs_error()
%>
        </td>
    </tr>
</table>
<%
end sub

sub news()
%>
<form action="admin_boardset.asp?action=new" method=post name=FORM>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">发布版面： </div>
      </td>
      <td width="80%" class=tablebody1> 
        <%if master then%>
        <%
   sql="select boardid,boardtype from board"
   rs.open sql,conn,1,1
%>
        <select name="boardid" size="1">
<!--           <option value="0">论坛首页</option> -->
          <%
	do while not rs.eof
        response.write "<option value='"+CStr(rs("BoardID"))+"'>"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
%>
        </select>
        <%else%>
        <%
	sql="select boardtype from board where boardid="&boardid
	rs.open sql,conn,1,1
	boardtype=rs("boardtype")
%>
        <select name="boardid" size="1">
          <option value="<%=boardid%>"><%=boardtype%></option>
        </select>
        <%end if%>
      </td>
    </tr>
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
        <input type=text name=title size=60>
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
        <input type=Submit value="发 送" name=Submit">
        &nbsp; 
        <input type="reset" name="Clear" value="清 除">
      </td>
    </tr>
  </table>
</form>
<%
end sub

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

sub manage()
dim caneditann
caneditann=false
if (master or superboardmaster or boardmaster) and Cint(GroupSetting(26))=1 then
	caneditann=true
else
	caneditann=false
end if
	if   not caneditann  then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
	end if
	sql="select * from bbsnews where boardid<>0"
	rs.open sql,conn,1,1
%>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
		  <tr><th width="80%" valign=top height=22 >
标题
		  </th>
		  <th width="20%" >
操作
		  </th></tr>
<%do while not rs.eof%>
		  <tr><td width="80%" valign=top height=22 class=tablebody1><a href=admin_boardset.asp?action=edit&id=<%=rs("id")%>&boardid=<%=rs("boardid")%>><%=rs("title")%></a>
		  </td>
		  <td width="20%" class=tablebody2><a href=admin_boardset.asp?action=del&id=<%=rs("id")%>&boardid=<%=boardid%>>删除</a>
		  </td></tr>
<%
	rs.movenext
	loop
	rs.close
%></table>
<%
end sub

sub edit()
%>
<form action="admin_boardset.asp?action=updat&id=<%=request("id")%>" method=post>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
		  <tr><td  valign=top align=right class=tablebody1>发布版面：
		  </td>
		  <td  class=tablebody1>
<%
	dim sel
   	sql="select boardid,boardtype from board"
   	rs.open sql,conn,1,1
%>
<select name="boardid" size="1">
<option value="0" <%if request("boardid")=0 then%>selected<%end if%>>论坛首页</option>
<%
	do while not rs.eof
	if Clng(request("boardid"))=Clng(rs("boardid")) then
	sel="selected"
	else
	sel=""
	end if
        response.write "<option value='"+CStr(rs("BoardID"))+"' "&sel&">"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
%>        
          </select>
		  </td></tr>
<%
	sql="select * from bbsnews where id="&cstr(request("id"))
	rs.open sql,conn,1,1
%>
		  <tr><td width="20%" valign=top align=right class=tablebody1>
发布人：
		  </td>
		  <td width="80%"  class=tablebody1><input type=text name=username size=36 value=<%=rs("username")%>></td></tr>
		  <tr><td width="20%" valign=top align=right class=tablebody1>
标题：
		  </td>
		  <td width="80%"  class=tablebody1><input type=text name=title size=60 value=<%=rs("title")%>></td></tr>
		  <tr><td width="20%" valign=top align=right class=tablebody1>
内容：
		  </td>
		  <td width="80%" class=tablebody1><textarea cols=60 rows=6 name="content">
<%
	    content=replace(rs("content"),"<br>",chr(13))
        content=replace(content,"&nbsp;"," ")
            response.write ""&content&""
	    rs.close
%>
		  </textarea></td>
		  </tr>
		  <tr><td width="100%" valign=top colspan="2" align=center class=tablebody2>
<input type=Submit value="修 改" name=Submit"> &nbsp; <input type="reset" name="Clear" value="清 除">
		  </td></tr>
		</table>
</form>
<%
end sub

sub update()
	dim username,title,content
	dim caneditann
	caneditann=false
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(25))=1 then
	caneditann=true
	else
	caneditann=false
	end if
	if   not caneditann  then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
	end if
	if request("id")="" or not isnumeric(request("id"))  then
	Errmsg=Errmsg+"<br><li>请选择正确的公告。"
	founderr=true
	exit sub
	end if
	if request("boardid")<>"" or (not isInteger(request("boardid")))  then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br><li>请选择正确的论坛。"
		founderr=true
		exit sub
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
	else
		sql="select * from bbsnews where id="&cstr(request("id"))
		rs.open sql,conn,1,3
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
	end if
end sub

sub del()
	dim caneditann
	caneditann=false
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(26))=1 then
	caneditann=true
	else
	caneditann=false
	end if

	if   not caneditann  then
	Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
	founderr=true
	exit sub
	end if

	if request("id")="" and not isnumeric(request("id"))  then
	Errmsg=Errmsg+"<br><li>请选择正确的公告。"
	founderr=true
	exit sub
	end if
	
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	conn.execute("delete from bbsnews where id in ("&delid&")")
	myCache.name="AnnounceMents"&BoardID
	myCache.makeEmpty
	call success()
end sub

sub success()
%><br><br>
成功：新闻操作
<%
end sub

sub editbminfo()
dim master_1

%><form action ="admin_boardset.asp?action=saveditbm&boardid=<%=boardid%>" method=post> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="+CSTr(boardid)
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
   call dvbbs_error()
exit sub
   end if
if not master then
	if Board_Setting(33)=1  then
		master_1=split(rs("boardmaster"),"|")
		if membername<>master_1(0)  then
		Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。"
		call dvbbs_error()
		exit sub
		end if
	else
	Errmsg=Errmsg+"<br>"+"<li>您未有修改设置的权限。"
	call dvbbs_error()
	exit sub
	end if
end if
%>
             <input type='hidden' name=editid value='<%=boardid%>'>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
    <tr> 
      <th width="20%" height=22 class=tablebody2><b>字段名称：</b> </td>
      <th  > 
        <div align="center"><b>变量值：</b></div>
      </th>
    </th>
    <tr> 
      <td height=22 class=tablebody1  align="center">论坛名称：</td>
      <td  class=tablebody1>
	  <input type="text" name="boardtype" size="30" value='<%=htmlencode(rs("boardtype"))%>'>
	  </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2  align="center">版面说明：</td>
      <td  class=tablebody1> 
        <input type="text" name="readme" size="60" value='<%=htmlencode(rs("readme"))%>'>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody1  align="center">版主修改：</td>
      <td  class=tablebody1> 
        <input type="text" name="boardmaster" size="30" value='<%=rs("boardmaster")%>'>(多版主添加请用|分隔，如：沙滩小子|wodeail)
				<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2>&nbsp;</td>
      <td  class=tablebody2> 
        <input type="submit" name="Submit" value="提交">
      </td>
    </tr>
  </table>
</form>
<%
rs.close

end sub

sub savebminfo()
dim rname
dim readme,boardtype,boardmaster
	if request("boardid")<>"" or (not isInteger(request("boardid")))  then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br><li>请选择正确的论坛。"
		founderr=true
		exit sub
	end if
readme=checkStr(Request.form("readme"))
boardtype=checkStr(Request.form("boardtype"))
boardmaster=checkStr(Request.form("boardmaster"))
	if readme="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入论坛简介。"
		founderr=true
	end if
	if boardtype="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入论坛名称。"
		founderr=true
	end if
	if boardmaster="" then
		Errmsg=Errmsg+"<br>"+"<li>请输入管理成员。"
		founderr=true
	end if
	if founderr=true then 
	call dvbbs_error()
	exit sub
	end if
rname=split(boardmaster,"|")
for i=0 to ubound(rname)
sql="select top 1 username from [user] where username='"&replace(rname(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>论坛没有这个用户，不能添加为版主"
	exit sub
	exit for
end if
set rs=nothing
next
set rs=server.createobject("adodb.recordset")
sql = "select * from board where boardid="+Cstr(request("boardid"))
rs.open sql,conn,1,3
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
	rs("boardmaster") = boardmaster
	rs("readme") = readme
	rs("boardtype")=boardtype
	rs.Update 
	rs.Close 
if request("oldboardmaster")<>boardmaster then call addmaster(boardmaster,request("oldboardmaster"),1)
	response.write "<p>论坛修改成功！"
end sub

sub editbmset()
dim chkedit
dim master_1
chkedit=false
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
if isnull(rs(0)) or rs(0)="" then
Errmsg=Errmsg+"<br>"+"<li>本论坛还未有管理员。"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。 "
	call dvbbs_error()
else
%>
<form method="POST" action=admin_boardset.asp?action=savebmset&boardid=<%=request("boardid")%>>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
<tr> 
<td height="23" colspan="2" class=Tablebody2><b>论坛使用设置</b></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left><a name="setting1"><b>打开或者关闭论坛</b></a>[<a href="#top"><FONT color="#FFFFFF">顶部</FONT></a>]</th>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>版面语言</U></td>
<td width="50%" class=tablebody1> 
<select size=1 name="forum_setting(9)">
<option value="lang_gb.asp">简体中文
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>论坛当前状态</U><BR>维护期间可设置关闭论坛使用</td>
<td width="50%" class=tablebody1> 
<input type=radio name="Forum_Setting(21)" value=0 <%if cint(Forum_Setting(21))=0 then%>checked<%end if%>>打开&nbsp;
<input type=radio name="Forum_Setting(21)" value=1 <%if cint(Forum_Setting(21))=1 then%>checked<%end if%>>关闭&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>维护说明</U><BR>在论坛关闭情况下显示，支持html语法</td>
<td width="50%" class=tablebody1> 
<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>是否起用定时开关论坛</U></td>
<td width="50%" class=tablebody1> 
<input type=radio name="forum_setting(69)" value=0 <%if cint(Forum_Setting(69))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(69)" value=1 <%if cint(Forum_Setting(69))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>论坛开放时间</U><BR>请确认您已经设置起用定时开关功能<BR>起止小时用符号“|”分开</td>
<td width="50%" class=tablebody1> 
<input type=text size=10 name="forum_setting(70)" value="<%=forum_setting(70)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>首页模式</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(4)" value=0 <%if cint(Forum_Setting(4))=0 then%>checked<%end if%>>新版样式&nbsp;
<input type=radio name="Forum_Setting(4)" value=1 <%if cint(Forum_Setting(4))=1 then%>checked<%end if%>>旧版样式&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>导航菜单模式</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(71)" value=0 <%if cint(Forum_Setting(71))=0 then%>checked<%end if%>>顶部固定&nbsp;
<input type=radio name="Forum_Setting(71)" value=1 <%if cint(Forum_Setting(71))=1 then%>checked<%end if%>>左侧滑动&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>论坛基本信息</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛名称</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛的url</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛的建站日期</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="CreateDate" size="35" value="<%=Forum_info(12)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>主页名称</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>主页URL</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛首页Logo地址</U><BR>显示在论坛首页，添加论坛如果没有填写logo地址，则使用该内容</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛图片目录</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>发帖心情目录</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>发帖表情目录</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="emoturl" size="35" value="<%=Forum_info(10)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛头像目录</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="UserFaceurl" size="35" value="<%=Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>版权信息</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting4"></a><b>论坛联系资料</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛管理员Email</U><BR>给用户发送邮件时，显示的来源Email信息</td>
<td width="50%" class=tablebody1>  
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>悄悄话选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>新短消息弹出窗口</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(10)" value=0 <%if cint(Forum_Setting(10))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(10)" value=1 <%if cint(Forum_Setting(10))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting7"></a><b>论坛首页选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>首页显示论坛深度</U><BR>0代表一级，1代表2级，以此类推<BR>设置过大的论坛深度将影响论坛整体性能，请根据自己论坛情况做设置，建议设置为1</td>
<td width="50%" class=tablebody1> 
<input type=text size=10 name="forum_setting(5)" value="<%=forum_setting(5)%>"> 级
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>快速登录位置</U></td>
<td width="50%" class=tablebody1>  
<select name="Forum_Setting(31)">
<option value="1" <%if cint(Forum_Setting(31))=1 then%>selected<%end if%>>顶部
<option value="2" <%if cint(Forum_Setting(31))=2 then%>selected<%end if%>>底部
<option value="0" <%if cint(Forum_Setting(31))="0" then%>selected<%end if%>>不显示
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>是否显示过生日会员</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(29)" value=0 <%if cint(Forum_Setting(29))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(29)" value=1 <%if cint(Forum_Setting(29))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>是否显示周年纪念会员</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(43)" value=0 <%if cint(Forum_Setting(43))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(43)" value=1 <%if cint(Forum_Setting(43))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>首页是否显示点歌列表(停用)</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(12)" value=0 <%if cint(Forum_Setting(12))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(12)" value=1 <%if cint(Forum_Setting(12))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>首页是否显示回收站</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(13)" value=0 <%if cint(Forum_Setting(13))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(13)" value=1 <%if cint(Forum_Setting(13))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>首页是否显示明星</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(16)" value=0 <%if cint(Forum_Setting(16))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(16)" value=1 <%if cint(Forum_Setting(16))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>明星第一列的显示方式</U></td>
<td width="50%" class=tablebody1><select name="Forum_Setting(17)"> 
<option value=1 <%if cint(Forum_Setting(17))=1 then%>selected<%end if%>>今日发帖量</option>
<option value=2 <%if cint(Forum_Setting(17))=2 then%>selected<%end if%>>本周发帖量</option>
<option value=3 <%if cint(Forum_Setting(17))=3 then%>selected<%end if%>>本月发帖量</option>
<option value=4 <%if cint(Forum_Setting(17))=4 then%>selected<%end if%>>今年发帖量</option>
<option value=5 <%if cint(Forum_Setting(17))=5 then%>selected<%end if%>>发帖总数</option>
<option value=6 <%if cint(Forum_Setting(17))=6 then%>selected<%end if%>>最佳男明星</option>
<option value=7 <%if cint(Forum_Setting(17))=7 then%>selected<%end if%>>最佳女明星</option>
<option value=8 <%if cint(Forum_Setting(17))=8 then%>selected<%end if%>>今日得分量</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>明星第二列的显示方式</U></td>
<td width="50%" class=tablebody1><select name="Forum_Setting(18)"> 
<option value=1 <%if cint(Forum_Setting(18))=1 then%>selected<%end if%>>今日发帖量</option>
<option value=2 <%if cint(Forum_Setting(18))=2 then%>selected<%end if%>>本周发帖量</option>
<option value=3 <%if cint(Forum_Setting(18))=3 then%>selected<%end if%>>本月发帖量</option>
<option value=4 <%if cint(Forum_Setting(18))=4 then%>selected<%end if%>>今年发帖量</option>
<option value=5 <%if cint(Forum_Setting(18))=5 then%>selected<%end if%>>发帖总数</option>
<option value=6 <%if cint(Forum_Setting(18))=6 then%>selected<%end if%>>最佳男明星</option>
<option value=7 <%if cint(Forum_Setting(18))=7 then%>selected<%end if%>>最佳女明星</option>
<option value=8 <%if cint(Forum_Setting(18))=8 then%>selected<%end if%>>今日得分量</option>
</select>
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>用户与注册选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>是否允许新用户注册</U><BR>关闭后论坛将不能注册</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(37)" value=0 <%if cint(forum_setting(37))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if cint(forum_setting(37))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>最短用户名长度</U><BR>填写数字，不能小于1大于50</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(40)" size="3" value="<%=forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>最长用户名长度</U><BR>填写数字，不能小于1大于50</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(41)" size="3" value="<%=forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>同一IP注册间隔时间</U><BR>如不想限制可填写0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(22)" size="3" value="<%=Forum_Setting(22)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>Email通知密码</U><BR>确认您的站点支持发送mail，所包含密码为系统随机生成</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(23)" value=0 <%if cint(Forum_Setting(23))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(23)" value=1 <%if cint(Forum_Setting(23))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>一个Email只能注册一个帐号</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(24)" value=0 <%if cint(Forum_Setting(24))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(24)" value=1 <%if cint(Forum_Setting(24))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>注册需要管理员认证</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(25)" value=0 <%if cint(Forum_Setting(25))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(25)" value=1 <%if cint(Forum_Setting(25))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>发送注册信息邮件</U><BR>请确认您打开了邮件功能</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(47)" value=0 <%if cint(Forum_Setting(47))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(47)" value=1 <%if cint(Forum_Setting(47))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>开启短信欢迎新注册用户</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(46)" value=0 <%if cint(Forum_Setting(46))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(46)" value=1 <%if cint(Forum_Setting(46))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>系统设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>论坛所在时区</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>服务器时差</U></td>
<td width="50%" class=tablebody1>  
<select name="Forum_Setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if cint(i)=cint(Forum_Setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>脚本超时时间</U><BR>默认为300，一般不做更改</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(1)" size="3" value="<%=Forum_Setting(1)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>是否显示页面执行时间</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(30)" value=0 <%if cint(Forum_Setting(30))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(30)" value=1 <%if cint(Forum_Setting(30))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>禁止的邮件地址</U><BR>在下面指定的邮件地址将被禁止注册，每个邮件地址用“|”符号分隔<BR>本功能支持模糊搜索，如设置了eway禁止，将禁止eway@aspsky.net或者eway@dvbbs.net类似这样的注册</td>
<td width="50%" class=tablebody1> 
<input type="text" name="Forum_Setting(52)" size="50" value="<%=Forum_Setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>防刷新功能有效的页面</U><BR>请确认您打开了防刷新功能<BR>您指定的页面将有防刷新作用，用户在限定的时间内不能重复打开该页面，具有一定减少资源消耗的作用<BR>每个页面名请用“|”符号隔开</td>
<td width="50%" class=tablebody1> 
<input type="text" name="Forum_Setting(64)" size="50" value="<%=Forum_Setting(64)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>在线和用户来源</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线显示用户IP</U><BR>关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(28)" value=0 <%if cint(Forum_Setting(28))=0 then%>checked<%end if%>>保密&nbsp;
<input type=radio name="Forum_Setting(28)" value=1 <%if cint(Forum_Setting(28))=1 then%>checked<%end if%>>公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线显示用户来源</U><BR>关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见<BR>开启本功能较消耗资源</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(36)" value=0 <%if cint(forum_setting(36))=0 then%>checked<%end if%>>保密&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if cint(forum_setting(36))=1 then%>checked<%end if%>>公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线资料列表显示用户当前位置</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(33)" value=0 <%if cint(forum_setting(33))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if cint(forum_setting(33))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线资料列表显示用户登录和活动时间</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(34)" value=0 <%if cint(forum_setting(34))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if cint(forum_setting(34))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线资料列表显示用户浏览器和操作系统</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(35)" value=0 <%if cint(forum_setting(35))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(35)" value=1 <%if cint(forum_setting(35))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线名单显示客人在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(15)" value=0 <%if cint(Forum_Setting(15))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(15)" value=1 <%if cint(Forum_Setting(15))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>在线名单显示用户在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(14)" value=0 <%if cint(Forum_Setting(14))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(14)" value=1 <%if cint(Forum_Setting(14))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>删除不活动用户时间</U><BR>可设置删除多少分钟内不活动用户<BR>单位：分钟，请输入数字</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(8)" size="3" value="<%=Forum_Setting(8)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>总论坛允许同时在线数</U><BR>如不想限制，可设置为0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(26)" size="6" value="<%=Forum_Setting(26)%>">&nbsp;人
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting13"></a><b>邮件选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>发送邮件组件</U><BR>如果您的服务器不支持下列组件，请选择不支持</td>
<td width="50%" class=tablebody1>  
<select name="Forum_Setting(2)">
<option value="0" <%if Forum_Setting(2)=0 then%>selected<%end if%>>不支持 
<option value="1" <%if Forum_Setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Forum_Setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Forum_Setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>SMTP Server地址</U><BR>只有在论坛使用设置中打开了发送邮件功能，该填写内容方有效</td>
<td width="50%" class=tablebody1>  
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting14"></a><b>上传设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>头像上传</U></td>
<td width="50%" class=tablebody1> 
<input type=radio name="Forum_Setting(7)" value=0 <%if Forum_Setting(7)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(7)" value=1 <%if Forum_Setting(7)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>允许的最大头像文件大小</U></td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(56)" size="6" value="<%=forum_setting(56)%>">&nbsp;K
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting15"></a><b>用户选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>允许个人签名</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(42)" value=0 <%if forum_setting(42)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(42)" value=1 <%if forum_setting(42)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>允许用户使用头像</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(53)" value=0 <%if forum_setting(53)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(53)" value=1 <%if forum_setting(53)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>最大头像宽度</U><BR>定义内容为头像的最大宽度</td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(57)" size="6" value="<%=forum_setting(57)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>最大头像高度</U><BR>定义内容为头像的最大高度</td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(58)" size="6" value="<%=forum_setting(58)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>默认头像宽度</U><BR>定义内容为论坛头像的默认宽度</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(38)" size="6" value="<%=forum_setting(38)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>默认头像高度</U><BR>定义内容为论坛头像的默认宽度</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(39)" size="6" value="<%=forum_setting(39)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>使用自定义头像的最少发帖数</U></td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(54)" size="6" value="<%=forum_setting(54)%>">&nbsp;篇
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>允许从其他站点上传头像</U><BR>就是是否可以直接使用http..这样的url来直接显示头像</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(55)" value=0 <%if forum_setting(55)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(55)" value=1 <%if forum_setting(55)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>用户签名是否开启UBB代码</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(65)" value=0 <%if Forum_Setting(65)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(65)" value=1 <%if Forum_Setting(65)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>用户签名是否开启HTML代码</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(66)" value=0 <%if Forum_Setting(66)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(66)" value=1 <%if Forum_Setting(66)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>用户是否开启帖图标签</U><BR>包括图片、flash、多媒体等</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(67)" value=0 <%if Forum_Setting(67)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(67)" value=1 <%if Forum_Setting(67)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>用户排行列表个数</U></td>
<td class=tablebody1 width="50%"> 
<input type="text" name="Forum_Setting(68)" size="6" value="<%=Forum_Setting(68)%>">&nbsp;个
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>用户头衔</U><BR>是否允许用户自定义头衔</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(6)" value=0 <%if cint(Forum_Setting(6))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(6)" value=1 <%if cint(Forum_Setting(6))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>用户头衔最大长度</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(59)" size="6" value="<%=forum_setting(59)%>">&nbsp;byte
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>自定义头衔最少发帖数量限制</U><BR>不做限制请设置为0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(60)" size="6" value="<%=forum_setting(60)%>">&nbsp;篇
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>自定义头衔注册天数限制</U><BR>不做限制请设置为0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(61)" size="6" value="<%=forum_setting(61)%>">&nbsp;天
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>自定义头衔上面两个条件加在一起限制</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(62)" value=0 <%if cint(Forum_Setting(62))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(62)" value=1 <%if cint(Forum_Setting(62))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>自定义头衔中要屏蔽的词语</U><BR>每个限制字符用“|”符号隔开</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(63)" size="50" value="<%=forum_setting(63)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting17"></a><b>防刷新机制</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>防刷新机制</U><BR>如选择打开请填写下面的限制刷新时间<BR>对版主和管理员无效</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(19)" value=0 <%if cint(Forum_Setting(19))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(19)" value=1 <%if cint(Forum_Setting(19))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>浏览刷新时间间隔</U><BR>填写该项目请确认您打开了防刷新机制<BR>仅对帖子列表和显示帖子页面起作用</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(20)" size="3" value="<%=Forum_Setting(20)%>">&nbsp;秒
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting18"></a><b>论坛分页设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>每页显示最多纪录</U><BR>用于论坛所有和分页有关的项目（帖子列表和浏览帖子除外）</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(11)" size="3" value="<%=Forum_Setting(11)%>">&nbsp;条
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting16"></a><b>帖子选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>作为热门话题的最低人气值</U><BR>标准为主题回复数</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(44)" size="3" value="<%=forum_setting(44)%>">&nbsp;条
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>编辑过的帖子显示"由xxx于yyy编辑"的信息</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(48)" value=0 <%if cint(forum_setting(48))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(48)" value=1 <%if cint(forum_setting(48))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>管理员编辑后显示"由XXX编辑"的信息</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(49)" value=0 <%if cint(forum_setting(49))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(49)" value=1 <%if cint(forum_setting(49))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>等待"由XXX编辑"信息显示的时间</U><BR>允许用户编辑自己的帖子而不在帖子底部显示"由XXX编辑"信息的时限(以分钟为单位)</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(50)" size="3" value="<%=forum_setting(50)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>编辑帖子时限</U><BR>编辑处理帖子的时间限制(以分钟为单位, 1天是1440分钟) 超过这个时间限制, 只有管理员和版主才能编辑和删除帖子. 如果不想使用这项功能, 请设置为0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(51)" size="3" value="<%=forum_setting(51)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>是否起用记录帖子阅读用户功能</U><BR>只有用户发言时选择记录阅读该帖用户方生效<BR>开启本功能将消耗部分系统资源<BR>发帖用户、管理员和版主可看到该帖阅读用户列表</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(3)" value=0 <%if cint(forum_setting(3))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(3)" value=1 <%if cint(forum_setting(3))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>帖子中广告模式</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(27)" value=0 <%if cint(Forum_Setting(27))=0 then%>checked<%end if%>>静态广告&nbsp;
<input type=radio name="Forum_Setting(27)" value=1 <%if cint(Forum_Setting(27))=1 then%>checked<%end if%>>滚动新帖&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>帖子中会员属性显示方式</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(72)" value=0 <%if cint(Forum_Setting(72))=0 then%>checked<%end if%>>形象详细&nbsp;
<input type=radio name="Forum_Setting(72)" value=1 <%if cint(Forum_Setting(72))=1 then%>checked<%end if%>>形象简单&nbsp;
<input type=radio name="Forum_Setting(72)" value=2 <%if cint(Forum_Setting(72))=2 then%>checked<%end if%>>头像详细&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>帖子中是否显示存款和股票</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(45)" value=0 <%if cint(Forum_Setting(45))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(45)" value=1 <%if cint(Forum_Setting(45))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>门派设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>是否开启论坛门派</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(32)" value=0 <%if cint(Forum_Setting(32))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(32)" value=1 <%if cint(Forum_Setting(32))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Tablebody1> &nbsp;</td>
<td width="50%" class=Tablebody1>  
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end if
end sub

'====================保存设置===================
sub savebmset()
dim chkedit
dim master_1
dim Forum_copyright
chkedit=false
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
if isnull(rs(0)) then
Errmsg=Errmsg+"<br>"+"<li>本论坛还未有管理员。"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
dim Forum_setting
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。 "
	call dvbbs_error()
else
Forum_Setting=request.Form("Forum_Setting(0)") & "," & request.Form("Forum_Setting(1)") & "," & request.Form("Forum_Setting(2)") & "," & request.Form("Forum_Setting(3)") & "," & request.Form("Forum_Setting(4)") & "," & request.Form("Forum_Setting(5)") & "," & request.Form("Forum_Setting(6)") & "," & request.Form("Forum_Setting(7)") & "," & request.Form("Forum_Setting(8)") & "," & request.Form("Forum_Setting(9)") & "," & request.Form("Forum_Setting(10)") & "," & request.Form("Forum_Setting(11)") & "," & request.Form("Forum_Setting(12)") & "," & request.Form("Forum_Setting(13)") & "," & request.Form("Forum_Setting(14)") & "," & request.Form("Forum_Setting(15)") & "," & request.Form("Forum_Setting(16)") & "," & request.Form("Forum_Setting(17)") & "," & request.Form("Forum_Setting(18)") & "," & request.Form("Forum_Setting(19)") & "," & request.Form("Forum_Setting(20)") & "," & request.Form("Forum_Setting(21)") & "," & request.Form("Forum_Setting(22)") & "," & request.Form("Forum_Setting(23)") & "," & request.Form("Forum_Setting(24)") & "," & request.Form("Forum_Setting(25)") & "," & request.Form("Forum_Setting(26)") & "," & request.Form("Forum_Setting(27)") & "," & request.Form("Forum_Setting(28)") & "," & request.Form("Forum_Setting(29)") & "," & request.Form("Forum_Setting(30)") & "," & request.Form("Forum_Setting(31)") & "," & request.Form("Forum_Setting(32)") & "," & request.Form("Forum_Setting(33)") & "," & request.Form("Forum_Setting(34)") & "," & request.Form("Forum_Setting(35)") & "," & request.Form("Forum_Setting(36)") & "," & request.Form("Forum_Setting(37)") & "," & request.Form("Forum_Setting(38)") & "," & request.Form("Forum_Setting(39)") & "," & request.Form("Forum_Setting(40)") & "," & request.Form("Forum_Setting(41)") & "," & request.Form("Forum_Setting(42)") & "," & request.Form("Forum_Setting(43)") & "," & request.Form("Forum_Setting(44)") & "," & request.Form("Forum_Setting(45)") & "," & request.Form("Forum_Setting(46)") & "," & request.Form("Forum_Setting(47)") & "," & request.Form("Forum_Setting(48)") & "," & request.Form("Forum_Setting(49)") & "," & request.Form("Forum_Setting(50)") & "," & request.Form("Forum_Setting(51)") & "," & request.Form("Forum_Setting(52)") & "," & request.Form("Forum_Setting(53)") & "," & request.Form("Forum_Setting(54)") & "," & request.Form("Forum_Setting(55)") & "," & request.Form("Forum_Setting(56)") & "," & request.Form("Forum_Setting(57)") & "," & request.Form("Forum_Setting(58)") & "," & request.Form("Forum_Setting(59)") & "," & request.Form("Forum_Setting(60)") & "," & request.Form("Forum_Setting(61)") & "," & request.Form("Forum_Setting(62)") & "," & request.Form("Forum_Setting(63)") & "," & request.Form("Forum_Setting(64)") & "," & request.Form("Forum_Setting(65)") & "," & request.Form("Forum_Setting(66)") & "," & request.Form("Forum_Setting(67)") & "," & request.Form("Forum_Setting(68)") & "," & request.Form("Forum_Setting(69)") & "," & request.Form("Forum_Setting(70)") & "," & request.Form("Forum_Setting(71)") & "," & request.Form("Forum_Setting(72)")


Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT") & "," & Request("emoturl") & "," & Request("userfaceurl") & "," & Request("CreateDate")
Forum_copyright=request("copyright")
sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"' where id="&skinid
conn.execute(sql)
response.write "设置成功。"
end if
end sub

sub editbmcolor()
dim master_1,chkedit
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
if isnull(rs(0)) then
Errmsg=Errmsg+"<br>"+"<li>本论坛还未有管理员。"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。 "
	call dvbbs_error()
else
%>
<form method="POST" action="?action=savebmcolor&boardid=<%=request("boardid")%>">
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
<script>
function color(para_URL){var URL =new String(para_URL)
window.open(URL,'','width=300,height=220,noscrollbars')}
</SCRIPT>
<tr> 
<th height="23" colspan="3" align=left>论坛界面设置&nbsp;单击 <a href="javascript:color('color.asp')"><font color=#FFFFFF>这里</font></a> 使用万用颜色拾取器（修改以下设置必须具备一定网页知识）</th>
</tr>
<tr> 
<td width="45%" class=tablebody1>论坛BODY标签<br>
控制整个论坛风格的背景颜色或者背景图片等</td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<input type="text" name="ForumBody" size="50" value="<%=Forum_body(11)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>Body的CSS定义<BR>可定义内容为网页字体颜色、背景、浏览器边框等<BR><FONT COLOR="red">以下定义中，凡是CSS定义的就不一定局限于颜色上的配置了，您可以设置其中的字体、背景等</FONT></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="iebarcolor" cols="50" rows="3"><%=Forum_body(25)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>论坛连接总的CSS定义<BR>可定义内容为连接字体颜色、样式等<BR></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="ForumCssLink" cols="50" rows="3"><%=Forum_body(6)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>论坛表格总的CSS定义<BR>这里为论坛总的表格定义，为一般表格的风格设置（如页面中除了导航和信息列表外的部分：尾部、头部等），可定义内容为背景、字体颜色、样式等<BR></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="ForumCssTD" cols="50" rows="3"><%=Forum_body(10)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>顶部菜单表格CSS定义(Logo & Banner上方)<BR><FONT COLOR="#000099">调用：Class=TopDarkNav</FONT></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="NavDarkcolor" cols="50" rows="3"><%=Forum_body(24)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>顶部菜单表格CSS定义(Logo & Banner下方)<BR><FONT COLOR="#000099">调用：Class=TopLighNav</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Navlighcolor" cols="50" rows="3"><%=Forum_body(23)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>顶部菜单表格CSS定义(导航菜单)<BR><FONT COLOR="#000099">调用：Class=TopLighNav1</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Navlighcolor1" cols="50" rows="3"><%=Forum_body(26)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>顶部菜单表格CSS定义(Logo & banner部分)<BR><FONT COLOR="#000099">调用：Class=TopLighNav2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Navlighcolor2" cols="50" rows="3"><%=Forum_body(28)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>表格边框颜色定义<br>
在下面定义一和二CSS定义控制不到的地方，最好保持和边框CSS定义一中同样的颜色</td>
<td width="5%" bgcolor="<%=Forum_body(27)%>"></td>
<td width="50%" class=tablebody1> 
<input type=text name="btablebackcolor" value="<%=Forum_body(27)%>" size=35>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>表格边框（体）CSS定义一<br>
一般页面，可定义内容为主表格背景、背景图、边框、宽度等<BR><FONT COLOR="#000099">调用：Class=TableBorder1</font></td>
<td width="5%"  class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Tablebackcolor" cols="50" rows="3"><%=Forum_body(0)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>表格边框（体）CSS定义二<br>
分论坛导航表格边框及部分页面，可定义内容为主表格背景、背景图、边框、宽度等<BR><FONT COLOR="#000099">调用：Class=TableBorder2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="aTablebackcolor" cols="50" rows="3"><%=Forum_body(1)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>标题表格CSS定义一（深背景）<br>
一般为表格的标题状态栏颜色或者背景上的定义，可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：th</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Tabletitlecolor" cols="50" rows="3"><%=Forum_body(2)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>标题表格字体连接CSS定义（深背景）<br>
如果标题栏为深背景，在这里必须特别定义在标题栏中连接的颜色和字体的颜色一样，否则在这里将采用默认的连接颜色<BR>注意：这里的#TableTitleLink不能去掉，可以将几个项目分开写来达到连接的不同效果<BR><FONT COLOR="#000099">调用：id=TableTitleLink</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="TableTitleLink" cols="50" rows="3"><%=Forum_body(7)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>标题表格CSS定义二（浅背景）<br>
为二级标题状态栏颜色或者背景上的定义，如首页的论坛分类栏，可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：Class=TableTitle2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="aTabletitlecolor" cols="50" rows="3"><%=Forum_body(3)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>表格体CSS定义一<BR>可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：Class=TableBody1</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Tablebodycolor" cols="50" rows="3"><%=Forum_body(4)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>表格体颜色二(1和2颜色在显示中穿插)<BR>可定义内容为背景、背景图、字体及其颜色等<BR><FONT COLOR="#000099">调用：Class=TableBody2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="aTablebodycolor" cols="50" rows="3"><%=Forum_body(5)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>表格宽度</td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<input type="text" name="TableWidth" size="35" value="<%=Forum_body(12)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>警告提醒语句的颜色</td>
<td width="5%" bgcolor="<%=Forum_body(8)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="AlertFontColor" size="35" value="<%=Forum_body(8)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>显示帖子的时候，相关帖子，转发帖子，回复等的颜色</td>
<td width="5%" bgcolor="<%=Forum_body(9)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="ContentTitle" size="35" value="<%=Forum_body(9)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>首页连接颜色<BR>如版面连接</td>
<td width="5%" bgcolor="<%=Forum_body(14)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="BoardLinkColor" size="35" value="<%=Forum_body(14)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>一般用户名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(15)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="user_fc" size="35" value="<%=Forum_body(15)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>一般用户名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(16)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="user_mc" size="35" value="<%=Forum_body(16)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>版主名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(17)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="bmaster_fc" size="35" value="<%=Forum_body(17)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>版主名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(18)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="bmaster_mc" size="35" value="<%=Forum_body(18)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>管理员名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(19)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="master_fc" size="35" value="<%=Forum_body(19)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>管理员名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(20)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="master_mc" size="35" value="<%=Forum_body(20)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>贵宾名称字体颜色</td>
<td width="5%" bgcolor="<%=Forum_body(21)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="vip_fc" size="35" value="<%=Forum_body(21)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>贵宾名称上的光晕颜色</td>
<td width="5%" bgcolor="<%=Forum_body(22)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="vip_mc" size="35" value="<%=Forum_body(22)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>&nbsp;</td>
<td width="5%" class=tablebody2></td>
<td width="50%" class=tablebody2> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</table>
</form>
<%
end if
end sub
sub savebmcolor()
dim Forum_body,master_1
dim chkedit
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。 "
	call dvbbs_error()
else
Forum_body=request.form("Tablebackcolor") & "|||" & request.form("aTablebackcolor") & "|||" & request.form("Tabletitlecolor") & "|||" & request.form("aTabletitlecolor") & "|||" & request.form("Tablebodycolor") & "|||" & request.form("aTablebodycolor") & "|||" & request.form("ForumCssLink") & "|||" & request.form("TableTitleLink") & "|||" & request.form("AlertFontColor") & "|||" & request.form("ContentTitle") & "|||" & request.form("ForumCssTD") & "|||" & request.form("ForumBody") & "|||" & request.form("TableWidth") & "|||" & request.form("BodyFontColor") & "|||" & request.form("BoardLinkColor") & "|||" & request.form("user_fc") & "|||" & request.form("user_mc") & "|||" & request.form("bmaster_fc") & "|||" & request.form("bmaster_mc") & "|||" & request.form("master_fc") & "|||" & request.form("master_mc") & "|||" & request.form("vip_fc") & "|||" & request.form("vip_mc") & "|||" & request.form("NavLighColor") & "|||" & request.form("NavDarkColor") & "|||" & request.form("IEbarColor") & "|||" & request.form("Navlighcolor1") & "|||" & request.form("btablebackcolor") & "|||" & request.form("Navlighcolor2")
response.write skinid

sql = "update config set Forum_body='"&checkstr(Forum_body)&"' where id="&skinid
conn.execute(sql)

response.write "设置成功！"
end if
end sub

sub editbmads()
dim master_1,chkedit
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
if isnull(rs(0))	then
Errmsg=Errmsg+"<br>"+"<li>本论坛还未有管理员。"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。 "
	call dvbbs_error()
else
%>
<form method="POST" action="?action=savebmads&boardid=<%=request("boardid")%>">
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
<tr> 
<th height="23" colspan="2" class="tableHeaderText"><b>论坛广告设置</b>（如为设置分论坛，就是分论坛首页广告，下属页面为帖子显示页面）</th>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>首页顶部广告代码</B></font></td>
<td width="60%" class="Tablebody1"> 
<textarea name="index_ad_t" cols="50" rows="3"><%=Forum_ads(0)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>首页尾部广告代码</B></font></td>
<td width="60%" class="Tablebody1"> 
<textarea name="index_ad_f" cols="50" rows="3"><%=Forum_ads(1)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>开启首页浮动广告</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type=radio name="index_moveFlag" value=0 <%if Forum_ads(2)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="index_moveFlag" value=1 <%if Forum_ads(2)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页浮动广告图片地址</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="MovePic" size="35" value="<%=Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页浮动广告连接地址</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="MoveUrl" size="35" value="<%=Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页浮动广告图片宽度</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="move_w" size="3" value="<%=Forum_ads(5)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页浮动广告图片高度</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="move_h" size="3" value="<%=Forum_ads(6)%>">&nbsp;象素
</td>
</tr>
<input type=hidden name="Board_moveFlag" value=0>
<tr> 
<td width="40%" class="Tablebody1"><B>开启首页右下固定广告</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type=radio name="index_fixupFlag" value=0 <%if Forum_ads(13)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="index_fixupFlag" value=1 <%if Forum_ads(13)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页右下固定广告图片地址</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixupPic" size="35" value="<%=Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页右下固定广告连接地址</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixupUrl" size="35" value="<%=Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页右下固定广告图片宽度</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixup_w" size="3" value="<%=Forum_ads(10)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>论坛首页右下固定广告图片高度</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixup_h" size="3" value="<%=Forum_ads(11)%>">&nbsp;象素
</td>
</tr>
<input type=hidden name="Board_fixupFlag" value=0>
<tr> 
<td width="40%" class="Tablebody1">&nbsp;</td>
<td width="60%" class="Tablebody1"> 
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end if
end sub
sub savebmads()
dim master_1
dim chkedit
dim Forum_adsinfo
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>您没有指定相应论坛ID，不能进行管理。"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>本项功能为主版主专用。 "
	call dvbbs_error()
else

Forum_adsinfo=request("index_ad_t") & "$" & request("index_ad_f") & "$" & request("index_moveFlag") & "$" & request("MovePic") & "$" & request("MoveUrl") & "$" & request("move_w") & "$" & request("move_h") & "$" & request("Board_moveFlag") & "$" & request("fixupPic") & "$" & request("FixupUrl") & "$" & request("Fixup_w") & "$" & request("Fixup_h") & "$" & request("Board_fixupFlag") & "$" & request("index_fixupFlag")
'response.write Forum_ads
sql = "update config set Forum_ads='"&CheckStr(Forum_adsinfo)&"' where id="&skinid
conn.execute(sql)
	response.write skinid&"设置成功。"
end if
end sub

sub addmaster(s,o,n)
dim arr,pw,oarr
dim classname,titlepic
set rs=conn.execute("select title from usergroups where usergroupid=3")
classname=rs(0)
set rs=conn.execute("select titlepic from usertitle where usergroupid=3 order by Minarticle desc")
if not (rs.eof and rs.bof) then
titlepic=rs(0)
end if
randomize
pw=Cint(rnd*9000)+1000
arr=split(s,"|")
oarr=split(o,"|")
set rs=server.createobject("adodb.recordset")
for i=0 to Ubound(arr)
sql="select * from [user] where username='"& arr(i) &"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	rs.addnew
	rs("username")=arr(i)
	rs("userpassword")=md5(pw)
	rs("userclass")=classname
	rs("UserGroupID")=3
	rs("titlepic")=titlepic
	rs("userWealth")=100
	rs("userep")=30
	rs("usercp")=30
	rs("userisbest")=0
	rs("userdel")=0
	rs("userpower")=0
	rs("lockuser")=0
	rs.update
	str=str&"你添加了以下用户：<b>" &arr(i) &"</b> 密码：<b>"& pw &"</b><br><br>"
else
	if rs("UserGroupID")>3 then
		rs("userclass")=classname
		rs("UserGroupID")=3
		rs("titlepic")=titlepic
		rs.update
	end if
end if
rs.close
next

'判断原版主在其他版面是否还担任版主，如没有担任则撤换该用户职位
if n=1 then
dim iboardmaster
dim UserGrade,article
iboardmaster=false
for i=0 to ubound(oarr)
	set rs=conn.execute("select boardmaster from board")
	do while not rs.eof
		if instr("|"&trim(rs("boardmaster"))&"|","|"&trim(oarr(i))&"|")>0 then
			iboardmaster=true
			exit do
		end if
	rs.movenext
	loop
	if not iboardmaster then
	set rs=conn.execute("select userid,UserGroupID,article from [user] where username='"&trim(oarr(i))&"'")
	if not (rs.eof and rs.bof) then
		if rs(1)>2 then
		if not isnumeric(rs(2)) then article=0
		set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where Minarticle<="&rs(2)&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
		if not (UserGrade.eof and UserGrade.bof) then
			conn.execute("update [user] set UserGroupID="&UserGrade(2)&",titlepic='"&UserGrade(1)&"',userclass='"&UserGrade(0)&"' where userid="&rs(0))
		end if
		end if
	end if
	end if
	iboardmaster=false
next
end if
set rs=nothing
end sub
%>

