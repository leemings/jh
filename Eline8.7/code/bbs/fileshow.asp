<!--#include FILE="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!-- #include file="inc/chkinput.asp" -->
<%

dim username
dim abgcolor
dim bbsurl
bbsurl=getservepath(request.ServerVariables("server_name")&request.ServerVariables("URL"))
stats="论坛文件展示"
call nav()

if Cint(GroupSetting(49))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有浏览本论坛展区的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if

if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	if  request("boardid")<>"" and  isnumeric(request("boardid"))  then
	if  clng(request("boardid"))<>0 then
	call head_var(1,BoardDepth,0,0)
	else
	call head_var(2,0,"","")
	end if
	else
	call head_var(2,0,"","")
	end if
	if request("action")="send" then
	call card()
	elseif request("action")="save" then
	call cardsave()
	elseif request("action")="cards" then
	call showcard()
	else
	call main()
	end if
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
dim sid
dim F_ID, F_AnnounceID, F_BoardID, F_UserID ,F_Username, F_Filename, F_FileType, F_Type, F_FileSize, F_Readme, F_DownNum, F_ViewNum, F_DownUser, F_Flag, F_AddTime
dim F_typename
dim golist,showfile

if request("id")="" or not isnumeric(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>您还未有选择展示的文件。"
	founderr=true
			exit sub
else
	sid=clng(request("id"))
end if
%>
  <table cellpadding=3 cellspacing=1 align=center class=tableborder1  >
    <tr>
      <th  height=25 align=center><%=stats%></th>
    </tr>
<%
set rs=conn.execute("select  F_ID, F_AnnounceID, F_BoardID, F_UserID ,F_Username, F_Filename, F_FileType, F_Type ,F_FileSize, F_Readme, F_DownNum, F_ViewNum, F_DownUser, F_Flag, F_AddTime  from [DV_Upfile] where    F_ID="&sid)
if rs.eof and rs.bof then
response.write"<tr >"&_
						"<td height=26 class=tablebody1 align=center>所查看的文件已被删除或禁止展示！</td>"&_
						"</tr >"
else

conn.execute("update [DV_Upfile]  set F_ViewNum=F_ViewNum+1  where  F_ID="&sid)

F_ID=rs(0)
F_AnnounceID=rs(1)
F_BoardID=rs(2)
F_UserID=rs(3)
F_Username=trim(htmlencode(rs(4)))
F_Filename=trim(rs(5))
F_FileType=trim(rs(6))
F_Type=rs(7)
F_FileSize=rs(8)
if rs(9)<>"" or not isnull(rs(9)) then
F_Readme=Server.htmlencode(rs(9))
else
F_Readme="<font color=gray>未有注明</font></a>"
end if
F_DownNum=rs(10)
F_ViewNum=rs(11)
F_DownUser=rs(12)
F_Flag=rs(13)
F_AddTime =rs(14)

	if instr(F_Filename,"/")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
	F_Filename=bbsurl&"UploadFile/"&F_Filename
	end if

	if not isnull(F_AnnounceID) then
		F_AnnounceID=split(F_AnnounceID,"|")
		golist="<a href=""dispbbs.asp?Boardid="&F_BoardID&"&ID="&F_AnnounceID(0)&"&replyID="&F_AnnounceID(1)&"&skin=1"" target=""_blank"" title=""浏览相关帖子"">论坛相关帖子......［参与评论］</a>"
		else
		golist="暂没有相关帖子"
	end if

	select case F_Type
	case 1
	F_typename="图片集"
	showfile="<a onfocus=this.blur() href="""&htmlencode(FilterJS(F_Filename))&""" target=_blank><IMG SRC="""&htmlencode(FilterJS(F_Filename))&""" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>"
	'showfile="[IMG]"&htmlencode(FilterJS(F_Filename))&"[/IMG]"
	'showfile=dvbcode(showfile,UserGroupID,3)
	case 2
	F_typename="FLASH集"
	showfile="[flash=500,350]"&htmlencode(FilterJS(F_Filename))&"[/flash]"
	showfile=dvbcode(showfile,UserGroupID,3)
	case 3
	F_typename="音乐集"
	showfile="<img src=images/files/"&F_FileType&".gif border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""点击欣赏"">"&htmlencode(FilterJS(F_Filename))&"</a>"
	case 4
	F_typename="电影集"
	showfile="<img src=images/files/"&F_FileType&".gif border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""点击欣赏"">"&htmlencode(FilterJS(F_Filename))&"</a>"
	case else
	F_typename="文件集"
	showfile="[upload="&F_FileType&"]viewfile.asp?ID="&F_ID&"[/upload]"
	showfile=dvbcode(showfile,UserGroupID,3)
	end select
dim edit
edit=false

if GroupSetting(48)=1 then
if  master or superboardmaster or boardmaster  then
edit=true
elseif F_Username=membername then
edit=true
else
edit=false
end if
end if

response.write"<tr >"&_
      		"<td height=* class=tablebody1 align=center>感谢会员：<a href=dispuser.asp?name="&F_Username&">"&F_Username&"</a>　提供"&_
	  	"<hr size=1 width=80% >"&_
	  	"<br>"&showfile&"<br>"
if founduser then
response.write	"<table border=0 width=100%  cellspacing=0 cellpadding=0>"&_
		"<tr><td width=* align=left>"
	if edit=true then
	response.write"<a href=myfile.asp?action=edit&editid="&F_ID&" title=""编辑该文件"" ><img src="&Forum_info(7)&"editfile.gif  border=0 width=""10"" height=""10""></a>"
	response.write"&nbsp;<a href=myfile.asp?action=fdel&delid="&F_ID&" title=""删除该文件"" ><img src="&Forum_info(7)&"delete.gif  border=0 width=""10"" height=""10""></a>"
	end if
	response.write"</td><td width=20 align=right><a href=fileshow.asp?action=send&id="&F_ID&" ><img src="&Forum_info(7)&"newmail.gif		border=0 title=""发送给朋友""></a></td></tr></table>"
end if
%>
	  </td>
    </tr >
<tr><td class=tableBorder2 height=150 valign=top align=center >
 <table cellpadding=3 cellspacing=1 align=center class=tableborder2 width=100% align=center>
 <tr><td width=20% class=tablebody1 align=right>作者：</td><td width=80% class=tablebody1><a href=dispuser.asp?name=<%=F_Username%> ><%=F_Username%></a></td></tr>
 <tr><td width=20% class=tablebody2 align=right>类别：</td><td width=80% ><%=F_typename%></td></tr>
 <tr><td width=20% class=tablebody1 align=right>大小：</td><td width=80% class=tablebody1><%=F_FileSize%>　Byte</td></tr>
 <tr><td width=20% class=tablebody2 align=right>浏览：</td><td width=80% ><font color=<%=Forum_body(8)%>><%=F_ViewNum%></font>　次</td></tr>
<%if F_Type=0 then%>
 <tr><td width=20% class=tablebody1 align=right>下载：</td><td width=80% class=tablebody1><%=F_DownNum%>　次</td></tr>
<%end if%>
 <tr><td width=20% class=tablebody1 align=right>创建时间：</td><td width=80% class=tablebody1><%=F_AddTime%></td></tr>
 <tr><td width=20% class=tablebody2 align=right>相关帖子：</td><td width=80%><%=golist%></td></tr>
 <tr><td height=20 width=20% class=tablebody1  align=right>说明：</td><td width=80% class=tablebody1><%=F_Readme%></td></tr>
</table>
</td>
</tr>
<%
end if
rs.close
set rs=nothing
response.write"</table>"

end sub

'编写贺卡内容
sub card()

stats="编写贺卡"
dim sid,showfile
dim F_Filename,F_Type
dim frs
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>,不能编写贺卡。"
	founderr=true
	exit sub
end if

if request("id")="" or not isnumeric(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>您还未有选择展示的文件。"
	founderr=true
	exit sub
else
	sid=clng(request("id"))
end if

%>
<script src="inc/ubbcode.js"></script>
  <table cellpadding=3 cellspacing=1 align=center class=tableborder1  align=center>
    <tr>
      <th  height=25 align=center><%=stats%></th>
    </tr>
<%
set rs=conn.execute("select  F_ID,F_Username, F_Filename, F_FileType, F_Type, F_Readme, F_ViewNum, F_Flag  from [DV_Upfile] where    F_ID="&sid)
if rs.eof and rs.bof then
response.write"<tr >"&_
						"<td height=26 class=tablebody1 align=center>所查看的文件已被删除或禁止展示！</td>"&_
						"</tr >"
else
F_Filename=htmlencode(FilterJS(rs(2)))
	if instr(F_Filename,"/")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
	F_Filename=bbsurl&"UploadFile/"&F_Filename
	end if

F_Type=cint(rs(4))

select case F_Type
	case 1
	showfile="<a onfocus=this.blur() href="""&F_Filename&""" target=_blank><IMG SRC="""&F_Filename&""" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>"
	'showfile="[IMG]"&F_Filename&"[/IMG]"
	'showfile=dvbcode(showfile,UserGroupID,3)
	case 2
	showfile="[flash=500,350]"&F_Filename&"[/flash]"
	showfile=dvbcode(showfile,UserGroupID,3)
	case else
	Errmsg=Errmsg+"<br>"+"<li>该文件目前不支持发送。"
	founderr=true
		exit sub
end select
%>
<tr >
      <td height=* class=tablebody1 align=center>：：您要发送的文件：：
	  <hr size=1 width=80% >
	  <br><%=showfile%><br></td>
    </tr >
<tr>
<td class=tablebody1 align=center height=150 valign=top>
  <table cellpadding=3 cellspacing=1 align=center class=tableborder1  align=center>
   <form action="?action=save" method=post name=frmAnnounce >
       <tr>
      <td width="100"  align="right" class=tablebody2>主题：</td>
      <td width="*" class=tablebody1>
<input type="text" name="title" size="60"></td>
    </tr>
    <tr>
      <td width="100"  align="right" class=tablebody2>收卡人：</td>
      <td width="*" class=tablebody1>
<input type="text" name="rname" size="60">
<SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">选择</OPTION>
<%
set frs=server.createobject("adodb.recordset")
sql="select F_friend from Friend where F_username='"&membername&"' order by F_addtime desc"
frs.open sql,conn,1,1
do while not frs.eof
%>
			  <OPTION value="<%=frs(0)%>"><%=frs(0)%></OPTION> 
<%
frs.movenext
loop
end if
frs.close
set frs=nothing
%>
</SELECT>
</td>
    </tr>
	  <tr>
      <td width="100"  align="right" class=tablebody2>ＵＢＢ功能：</td>
      <td width="*" class=tablebody1><!--#include file="inc/getubb.asp"--></td>
    </tr>
  <tr>
    <td  height="128" class=tablebody2 align="right">附言：</td>
    <td  class=tablebody1>
<textarea rows="8"  cols="103" name=Content></textarea></td>
    </tr>
    <tr>
      <th  height="26" colspan="2" align=center>
<input type="submit" value="提交" >
<input type="reset" value="全部重写" >
<input type=hidden name="sname" value="<%=htmlencode(membername)%>">
 <input type=hidden name="saveid" value="<%= rs("F_ID") %>">
      </th>
    </tr></form>
  </table>
</td></tr>
<%
rs.close
set rs=nothing
response.write"</table>"
end sub

'=====================贺卡演示====================
sub showcard()
stats="欣赏贺卡"

dim cid,msnid
dim sender,incept,body,title,sendtime
dim F_Filename,ftype,flag
dim showfile

if request("id")="" or not isnumeric(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>您还未有选择展示的文件。"
	founderr=true
	exit sub
else
	cid=clng(request("id"))
end if

if request("msmid")="" or not isnumeric(request("msmid")) then
	Errmsg=Errmsg+"<br>"+"<li>您还未有选择展示的文件。"
	founderr=true
	exit sub
else
	msnid=clng(request("msmid"))
end if

'取出短信内容
set rs=conn.execute("select  sender,incept,title,content,sendtime  from message where id="&msnid&" order by id desc")
if not (rs.eof and rs.bof) then
sender=htmlencode(trim(rs(0)))
incept=htmlencode(trim(rs(1)))
title=htmlencode(rs(2))
body=rs(3)
sendtime=rs(4)
else
	Errmsg=Errmsg+"<br>"+"<li>贺卡信息已不存在。"
	founderr=true
	exit sub
end if
set rs=nothing
'取出文件内容
set rs=conn.execute("select F_Filename,F_Type,F_Flag  from  [DV_Upfile] where F_ID="&cid&" order by F_ID desc")
if not (rs.eof and rs.bof) then
F_Filename=rs(0)
ftype=cint(rs(1))
flag=rs(2)
	if flag<>3 then
	Errmsg=Errmsg+"<br>"+"<li>您选择的并非贺卡文件。"
	founderr=true
	exit sub
	end if
else
	Errmsg=Errmsg+"<br>"+"<li>贺卡文件已不存在。"
	founderr=true
	exit sub
end if
set rs=nothing

	if instr(F_Filename,"/")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
	F_Filename=bbsurl&"UploadFile/"&F_Filename
	end if

	select case ftype
	case 1
	showfile="<a onfocus=this.blur() href="""&htmlencode(FilterJS(F_Filename))&""" target=_blank><IMG SRC="""&htmlencode(FilterJS(F_Filename))&""" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>"
	'showfile="[IMG]"&htmlencode(FilterJS(F_Filename))&"[/IMG]"
	'showfile=dvbcode(showfile,UserGroupID,3)
	case 2
	showfile="[flash=500,350]"&htmlencode(FilterJS(F_Filename))&"[/flash]"
	showfile=dvbcode(showfile,UserGroupID,3)
	case 3
	showfile="<img src=images/files/"&F_FileType&".gif border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""点击欣赏"">"&htmlencode(FilterJS(F_Filename))&"</a>"
	case 4
	F_typename="电影集"
	showfile="<img src=images/files/"&F_FileType&".gif border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""点击欣赏"">"&htmlencode(FilterJS(F_Filename))&"</a>"
	case else
	showfile="[upload="&F_FileType&"]viewfile.asp?ID="&F_ID&"[/upload]"
	showfile=dvbcode(showfile,UserGroupID,3)
	end select

%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder2  width="200" height="1"  style="word-break:break-all">
<tr>
      <th  height=25 align=center><%=stats%></th>
    </tr>
	<tr>
      <td width="100%" align=left class=tablebody1>
	  <font color=gray><img src="<%=forum_info(7)%>gift.gif"  border=0>　<%=sendtime%>，由 <b><%=sender%></b> 送给 <b><%=incept%></b>　的贺卡
	  <br>标题：<font  face="Comic Sans MS,helvetica" color="<%=Forum_body(8)%>"><%=title%></font>
	  </td>
    </tr>
    <tr>
      <td width="100%" height="172" align=center bgcolor="#FFFFFF">
		<table cellpadding=3 cellspacing=0 align=center   height="1"  style="word-break:break-all">
	<tr>
      <td width="100%" align=center class=tablebody1>
	  <br><%=showfile%><br></td>
    </tr>
	<tr><td align=left height="26" class=tablebody1><font size="+1" face="Comic Sans MS,helvetica" color=black>To：<font color="<%=Forum_body(8)%>" ><b><%=incept%></b></font>：</font></td></tr>
		<tr><td align=center class=tablebody1>
		<marquee style="FLOAT: none; MARGIN-LEFT: auto; MARGIN-RIGHT: auto" scrollamount=2 direction=up width=500 height=200 border=0 scrolldelay=100>
	  <%=dvbcode(body,UserGroupID,3)%></marquee>
		</td></tr>
<tr><td align=right height="26" class=tablebody1><font  face="Comic Sans MS,helvetica" color=black><b>From：<font color="<%=Forum_body(8)%>" ><%=sender%></b></font><br><%=sendtime%></font></td></tr>
		</table>
	  </td>
    </tr>
  </table>
<%

end sub

'============================保存贺卡信息=====================
sub cardsave()
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
	exit sub
end if

dim cid,sname,rname,ctitle,body
dim msmid,cardurl,msmbody
cid=checkStr(trim(request.form("saveid")))
sname=checkStr(trim(request.form("sname")))
rname=checkStr(trim(request.form("rname")))
ctitle=checkStr(trim(request.form("title")))
body=checkStr(trim(request.form("Content")))

if cid="" or not isnumeric(cid) then
	Errmsg=Errmsg+"<br>"+"<li>您还未有选择展示的文件。"
	founderr=true
	exit sub
end if
if not (isnull(session("lastpost")) or boardmaster or master or superboardmaster) then
	if DateDiff("s",session("lastpost"),Now())<10 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>本论坛限制发帖距离时间为10秒，请稍后再发。"
   		FoundErr=True
		exit sub
	end if
end if
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
	FoundErr=True
		exit sub
end if
if rname="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入接受卡人姓名。"
	FoundErr=True
		exit sub
	else
	rname=split(rname,",")
end if
if ctitle="" then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>贺卡的主题不应为空。"
		exit sub
elseif strLength(ctitle)>50 then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>贺卡主题长度不能超过50"
		exit sub
end if
if strLength(body)>15360 then
	ErrMsg=ErrMsg+"<Br>"+"<li>附言不得大于 15KB"
	FoundErr=true
		exit sub
end if
if body="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>没有填写附言。"
	FoundErr=true
		exit sub
end if

for i=0 to ubound(rname)
sql="select username from [user] where username='"&replace(rname(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>论坛没有这个用户，看看你的发送对象写对了吗？"
	exit sub
	exit for
end if
set rs=nothing
	'插入短信并获得ID
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&rname(i)&"','"&membername&"','"&ctitle&"','"&body&"',Now(),0,1)"
conn.execute(sql)

set rs=conn.execute("select top 1 id from message order by id desc")
msmid=rs(0)
set rs=nothing
cardurl=bbsurl&"fileshow.asp?action=cards&id="&cid&"&msmid="&msmid&""
cardurl="[URL="&cardurl&"]点击欣赏贺卡[/URL]"
msmbody=body+chr(13)+chr(13)+chr(10)+chr(10)+chr(10)+cardurl

conn.execute("update [message] set content='"&checkStr(msmbody)&"' where id="&msmid&" ")
conn.execute("update [DV_Upfile] set F_Flag=3  where F_ID="&cid&" ")

if i>Cint(GroupSetting(33))-1 then
errmsg=errmsg+"<br>"+"<li>最多只能发送给"&GroupSetting(33)&"个用户，您的名单"&GroupSetting(33)&"位以后的请重新发送"
	founderr=true
	exit sub
	exit for
end if
cardurl=""
next
sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，发送短信息成功。</b>。"
session("lastpost")=Now()
'response.write cid&"|"&sname&"|"&rname&"|"&ctitle&"|"&body
call dvbbs_suc()
end sub

function getservepath(str)
dim tmpstr
tmpstr=split(str,"/")
getservepath="http://"&replace(str, tmpstr(ubound(tmpstr)), "")
end function
%>
<script language="javascript"> 
function DoTitle(addTitle) {  
 var revisedTitle;  
 var currentTitle = document.frmAnnounce.rname.value; 

 if(currentTitle=="") revisedTitle = addTitle; 
 else { 
  var arr = currentTitle.split(","); 
  for (var i=0; i < arr.length; i++) { 
   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
  } 
  revisedTitle = currentTitle+","+addTitle; 
 } 

 document.frmAnnounce.rname.value=revisedTitle;  
 document.frmAnnounce.rname.focus(); 
 return; 
} 
</script>