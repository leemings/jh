<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
  <script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
 <table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
    <tr>
      <th  colspan="3" align="center"><a href=?><font color=#ffffff>论坛模版导出功能</font></a> | <a href=?action=load><font color=#ffffff>论坛模版导入功能</font></a></th>
    </tr>
  <tr >
    <td   align=center class="forumrow">
      <%
dim admin_flag
admin_flag="44"
if not master or instr(session("flag"),admin_flag)=0 then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
             if request("action")="inputskin" then
                        call inputskin()
                elseif request("action")="loadskin" then
                        call loadskin()
                elseif request("action")="load" then
                        call load()
                elseif request("action")="rename" then
                        call rename()
                elseif request("action")="savenm" then
                        call savenm()
                else
						CALL MAIN()
                end if
end if
%>
</td>
</tr>
    <tr >
      <th height="1" colspan="3" align="center"><a href=?><font color=#ffffff>论坛模版导出功能</font></a> | <a href=?action=load><font color=#ffffff>论坛模版导入功能</font></a></th>
    </tr>
  </table>

</body>
</html>
<%SUB MAIN()
dim sname,act,mdbname,tconn
if request("action")="loadthis" then
    sname="导入"
    act="loadskin"
    mdbname=Checkstr(trim(request.form("skinmdb")))
    if mdbname="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导出模版保存的表名"
   		response.write Errmsg
            exit sub
end if
else
    sname="导出"
    act="inputskin"
end if     

%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  <tr >
    <td width="100%" height="1" colspan="4" align="center" class="forumrow">
<B><%=sname%>论坛模版列表</B></td>
  </tr>
     <tr >
      <td width="10%"  align="center" class="forumrow">序号</td>
      <td width="70%"  align="center" class="forumrow">模版名称</td>
      <td width="10%"  align="center" class="forumrow">操作</td>
      <td width="10%"  align="center" class="forumrow">选择</td>
    </tr>
 <form action="?action=<%=act%>" method=post name=even>
<%
if act="loadskin" then
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	set rs=tconn.execute(" select id,Forum_Body,SkinName,active from dvskin  order by id ")
else
	set rs=conn.execute(" select id,Forum_Body,SkinName,active from config  order by id ")
end if
	do while not rs.eof
	response.write "<tr "
	if rs("active")=1 then
	response.write "bgcolor="&Forum_body(2)&" "
	else
	response.write "bgcolor="&Forum_body(2)&" "
	end if
	response.write ">"
%>
	<td class="forumrow"><%=rs("id")%></td>
	<td class="forumrow"><%=rs("SkinName")%></td>
	<td class="forumrow" align=center><a href="?action=rename&act=<%=act%>&skid=<%=rs("id")%>&mdbname=<%=mdbname%>" >改名</a></td>
	<td class="forumrow"><p><input type="checkbox" name="skid" value="<%=rs("id")%>"><input TYPE="hidden" NAME="active" VALUE="<% =rs("active") %>"></p></td>
	</tr>
<%		rs.movenext
		loop
     rs.close
     set rs=nothing
%>
  <tr >
    <td  colspan="4" align=center class="forumrow">
<%=sname%>的数据库：<input type="text" name="skinmdb" size="30" value="skin/skin.mdb">
<input type="submit" name="submit" value="<%=sname%>"> <input type=submit name=Submit value=删除  onclick="{if(confirm('注意：所删除的模版将不能恢复！')){this.document.even.submit();return true;}return false;}">  <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">全选</td>
  </tr>
</form>
  <tr >
    <td class="forumrow" colspan="4" align=left>
注意<br>
１，确认模版数据库名正确；<br>
２，如模版数据库放在skin目录下，即填写：skin/skin.mdb；<br>
３，模版数据库内备份的表名为DVSKIN,请不要更改；<br>
４，模版数据包括论坛ＣＳＳ设置，与及所有论坛图片设置．
</td>
  </tr>
</table>
<% END SUB 


sub inputskin()
dim skid,mdbname,TCONN
	skid=checkstr(request("skid"))
   	mdbname=Checkstr(trim(request.form("skinmdb")))
if skid="" or  isnull(skid) then
	ErrMsg=ErrMsg+"<Br>"+"<li>您还未选取要导出的模版"
	response.write Errmsg
	exit sub
end if
if mdbname="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请请填写导出模版数据库名"
	response.write Errmsg
	exit sub
end if
response.write "模版ＩＤ："&skid

if request("submit")="删除" then
	set rs=conn.execute("select skinname from config where active=1 and id in ("&skid&")")
	if rs.eof and rs.bof then
	conn.execute("delete from config where id in ("&skid&")")
	response.write "<br>删除成功。"
	else
	response.write "<br>错误提示：当前使用的论坛模板不能删除，请先设置该模板为不是当前模板状态。"
	end if
	rs.close
	set rs=nothing

else
    
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	dim trs,nrs
	set rs=conn.execute(" select Forum_Body,SkinName,Forum_pic,Forum_boardpic,Forum_TopicPic,Forum_statePic,Forum_ubb,Forum_userface,Forum_PostFace,Forum_emot from config where id in ("&skid&")  order by id ")
	do while not rs.eof
	tconn.execute("insert into dvskin (Forum_Body,SkinName,Forum_pic,Forum_boardpic,Forum_TopicPic,Forum_statePic,Forum_ubb,Forum_userface,Forum_PostFace,Forum_emot) values ('"&rs("Forum_Body")&"','"&rs("SkinName")&"','"&rs("Forum_pic")&"','"&rs("Forum_boardpic")&"','"&rs("Forum_TopicPic")&"','"&rs("Forum_statePic")&"','"&rs("Forum_ubb")&"','"&rs("Forum_userface")&"','"&rs("Forum_PostFace")&"','"&rs("Forum_emot")&"')")
	rs.movenext
	loop
	rs.close
	set rs=nothing
	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li> 数据库操作失败，请以后再试，"&err.Description
		err.clear
		response.write Errmsg
		exit sub
	else
		response.write "数据导出成功！<br>"
	end if
end if
end sub

sub load()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  <tr >
      <th  colspan="2">
        <p align="center">导入模版数据</th>
    </tr>
 <form action="?action=loadthis" method=post >
    <tr >
      <td width="20%" class="forumrow">导入模版数据库名：</td>
      <td width="80%" class="forumrow"><input type="text" name="skinmdb" size="30" value="skin/skin.mdb"></td>
    </tr>
    <tr >
      <th  colspan="2">
        <p align="center"><input type="submit" name="submit" value="下一步"></th>
    </tr>
</FORM>
  </table>
<%
end sub

sub loadskin()
dim skid,mdbname,TCONN

	skid=checkstr(request("skid"))
   	mdbname=Checkstr(trim(request.form("skinmdb")))
if skid="" or  isnull(skid) then
     		ErrMsg=ErrMsg+"<Br>"+"<li>您还未选取要导入的模版"
   		response.write Errmsg
            exit sub
end if
if mdbname="" then
     		ErrMsg=ErrMsg+"<Br>"+"<li>请填写导入模版数据库名"
   		response.write Errmsg
            exit sub
end if
response.write "模版ＩＤ："&skid

     dim trs,ars
       	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
   
if request("submit")="删除" then
	tconn.execute("delete from dvskin where id in ("&skid&")")
	response.write "删除成功。"

else

     set trs=tconn.execute(" select * from dvskin where id in ("&skid&")  order by id ")
		do while not trs.eof
     set Ars=conn.execute("select * from config where active=1")

     set rs = server.CreateObject ("adodb.recordset")
     sql="select * from config"
rs.open sql,conn,1,3
rs.addnew
'==========================导入的数据开始=======================
rs("skinname")=trs("skinname")
rs("Forum_body")=trs("Forum_body")
rs("Forum_pic")=trs("Forum_pic")
rs("Forum_boardpic")=trs("Forum_boardpic")
rs("Forum_TopicPic")=trs("Forum_TopicPic")
rs("Forum_statePic")=trs("Forum_statePic")
rs("Forum_ubb")=trs("Forum_ubb")
rs("Forum_userface")=trs("Forum_userface")
rs("Forum_PostFace")=trs("Forum_PostFace")
rs("Forum_emot")=trs("Forum_emot")
'==========================导入的数据结束=======================
rs("Forum_info")=ars("Forum_info")
rs("Forum_copyright")=ars("Forum_copyright")
rs("Forum_Setting")=ars("Forum_Setting")
rs("StopReadme")=ars("StopReadme")
rs("Forum_upload")=ars("Forum_upload")
rs("Forum_ads")=ars("Forum_ads")
rs("Forum_user")=ars("Forum_user")
rs("badwords")=ars("badwords")
rs("splitwords")=ars("splitwords")
rs("birthuser")=ars("birthuser")
rs("Maxonline")=ars("Maxonline")
rs("MaxonlineDate")=ars("MaxonlineDate")
rs("LastPost")="$$"&Now()&"$$$$$"
rs("allposttable")=ars("allposttable")
rs("allposttablename")=ars("allposttablename")
rs("NowUseBBS")=ars("NowUseBBS")
rs("NowUseBBS")=ars("NowUseBBS")
rs("Cookiepath")=ars("Cookiepath")
rs("Forum_sn")=ars("Forum_sn")
rs.update
rs.close
set rs=nothing

		trs.movenext
		loop
     trs.close
     set trs=nothing
	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li> 数据库操作失败，请以后再试，"&err.Description
		err.clear
		response.write Errmsg
		exit sub
	else
		response.write "数据导入成功！<br>"
	end if
call cache_skin()
end if

end sub

sub rename()
dim skid,mdbname,tconn
dim srs
	skid=checkstr(request("skid"))
        mdbname=Checkstr(trim(request("mdbname")))
if request("act")="loadskin" then      
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
 set srs=tconn.execute(" select id,Forum_Body,SkinName from dvskin where id="&skid&" ")
        else
 set srs=conn.execute(" select id,Forum_Body,SkinName from config where id="&skid&" ")
end if
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  <tr >
      <th   colspan="2">
        <p align="center">更改模版名称 ID=<%=srs("ID")%></td>
    </tr>
 <form action="?action=savenm" method=post >
    <tr >
      <td width="20%" class="forumrow">模版原名：</td>
      <td width="80%" class="forumrow"><%=srs("SkinName")%></td>
    </tr>
   <tr >
      <td width="20%" class="forumrow">模版新名：</td>
      <td width="80%" class="forumrow"><input type="text" name="skinNAME" size="30" value=""></td>
    </tr>
    <tr>
      <th   colspan="2">
        <p align="center"><input type="submit" name="submit" value="更新"></th>
    </tr>
<% if request("act")="loadskin" then
%><input TYPE="hidden" NAME="mdbname" VALUE="<% =mdbname %>">
<% end if %>
<input TYPE="hidden" NAME="skid" VALUE="<% =srs("id") %>">
<input TYPE="hidden" NAME="act" VALUE="<% =request("act") %>">
</FORM>
  </table>
<%

srs.close
set srs=nothing
end sub

sub savenm()
dim skid,mdbname,tconn,skinNAME
	skid=checkstr(request("skid"))
    mdbname=Checkstr(trim(request("mdbname")))
	skinNAME=Checkstr(trim(request.form("skinname")))
if request("act")="loadskin" then
        
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	tconn.execute(" UPDATE dvskin set SkinName='"&skinNAME&"'  where id="&skid&" ")
        else
	conn.execute(" UPDATE config set SkinName='"&skinNAME&"'  where id="&skid&" ")
end if
call cache_skin()
response.write "数据更新成功！<a href="&Request.ServerVariables("HTTP_REFERER")&" >点击返回</a><br>"

end SUB

sub cache_skin()
myCache.name="stylelist"
dim stylelist
stylelist="<a style=font-size:9pt;line-height:12pt; href=\""cookies.asp?action=stylemod&skinid=0&boardid={boardid}\"">恢复默认设置</a><br>"
set rs=conn.execute("select id,skinname from config order by id")
do while not rs.eof
stylelist=stylelist & "<a style=font-size:9pt;line-height:12pt; href=\""cookies.asp?action=stylemod&skinid="&rs(0)&"&boardid={boardid}\"">"&rs(1)&"</a><br>"
rs.movenext
loop
myCache.add stylelist,dateadd("n",60,now)
end sub
%>