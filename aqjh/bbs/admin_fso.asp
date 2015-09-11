<!-- #include file="setup.asp" -->




<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


dir=Request("dir")

response.write "<center>"
select case Request("menu")


case "ReForum"
ReForum


case "CreatTable"
Sql="CREATE TABLE ReForum"&Request("tablename")&" ("&_
	"id int IDENTITY (1, 1) NOT NULL ,"&_
	"topicid int NOT NULL ,"&_
	"username varchar(255) NOT NULL ,"&_
	"content text NOT NULL ,"&_
	"posttime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"postip varchar(255) NOT NULL "&_
")"
conn.execute(sql)

sql="insert into [TableList] (TableName) values ('"&Request("tablename")&"')"
conn.Execute(SQL)

error2("建立成功！")


case "update"
if Request("tablename") = "" Then ThisValue = "Null" Else ThisValue = Request("tablename")
conn.execute("update [clubconfig] set ReForumName="&ThisValue)
error2("设置成功！") 



error2("设置成功！")


case "del"

If not conn.Execute("Select ReList From [forum] where ReList="&request("tablename")&"" ).eof Then error2("该表中有对应的主题，不能删除！")
if ReForumName=request("tablename") then error2("当前正在使用中的数据库不能删除！")
conn.execute("delete from [TableList] where TableName="&request("tablename")&"")
conn.Execute("drop table ReForum"&request("tablename")&"")
error2("删除成功！")

case "clean"
maxForum=Request("maxForum")
if maxForum="" then maxForum=conn.execute("select Max(topicid) From [ReForum"&request("tablename")&"]")(0)
topicid=Request("topicid")
if topicid="" then topicid=0

sql="select top 200 topicid from [ReForum"&request("tablename")&"] where topicid > "&topicid&" order by topicid"
Set Rs=Conn.Execute(SQL)

Do While Not RS.EOF

if topicid<>rs("topicid") then
topicid=rs("topicid")

if conn.execute("Select id from [forum] where id="&topicid&"").eof Then conn.execute("delete from [ReForum"&request("tablename")&"] where topicid="&topicid&"")
i=i+1
end if

RS.MoveNext
loop
set rs=nothing

if topicid < maxForum then
plan=formatnumber(topicid/maxForum*100)
%>
<table border="0" width="60%">
	<tr>
		<td><img height="10" src="images/bar/0.gif" width="<%=plan%>%"> <%=plan%>%<br>
</td>
	</tr>
</table>

本次共执行 <%=i%> 次查询
<script>function jump(){
window.location.href='admin_fso.asp?menu=clean&tablename=<%=request("tablename")%>&maxForum=<%=maxForum%>&topicid=<%=topicid%>';
}
setTimeout(jump,2000);
</script>
<%
else
%>
已经将数据库[ReForum<%=request("tablename")%>]中的无用数据清理完成
<%
end if

case "files"
files

case "delfiles"

if dir<>"" then dir=""&dir&"/"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
for each ho in request.form("files")
MyFileObject.DeleteFile""&Server.MapPath("./images/upfile/"&dir&""&ho&"")&""
next
error2("已经成功删除所选的文件！")


case "bak"
bak

case "bakbf"
if instr(""&Request("bf")&"",".asp") > 0 or instr(""&Request("yl")&"",".asp") > 0 then error2("数据库路径中不能包含“.asp”")
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request("yl"))&"",""&Server.MapPath(Request("bf"))&""
error2("备份成功！")

case "bakhf"
if instr(""&Request("bf")&"",".asp") > 0 or instr(""&Request("yl")&"",".asp") > 0 then error2("数据库路径中不能包含“.asp”")
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request("bf"))&"",""&Server.MapPath(Request("yl"))&""
error2("恢复成功！")


case "statroom"
statroom

end select

sub bak
%>


<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>备份数据库</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
<form method="post" action="?menu=bakbf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%">数据库路径： </td>
<td width="70%">
<input size="30" value="<%=""&datapath&""&datafile&""%>" name="yl"></td>
</tr>
<tr>
<td width="30%">备份的数据库路径：</td>
<td width="70%"><input size="30" value="<%=replace(""&datapath&""&datafile&"",".mdb","bak.mdb")%>" name="bf"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" 备 份 " name="Submit1"><br></td>
</tr>
</table>
</td></tr></table></form>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>恢复数据库</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<form method="post" action="?menu=bakhf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%">备份的数据库路径： </td>
<td width="70%">
<input size="30" value="<%=replace(""&datapath&""&datafile&"",".mdb","bak.mdb")%>" name="bf"></td>
</tr>
<tr>
<td width="30%">数据库路径：</td>
<td width="70%"><input size="30" value="<%=""&datapath&""&datafile&""%>" name="yl"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" 恢 复 " name="Submit1"><br></td>
</tr>
</table></td></tr></table>
</form>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>压缩数据库</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>

<form action=compact.asp>
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>

<td width="70%">


<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>

<td width="30%">数据库路径： </td>
<td width="70%">
<input size="30" value="<%=""&datapath&""&datafile&""%>" name="dbpath"></td>
</tr>
<tr>
<td width="30%">数据库格式：</td>
<td width="70%"><input type="radio" value="True" name="boolIs97" id=boolIs97><label for=boolIs97>Access 97</label>　<input type="radio" value="" name="boolIs97" checked id=boolIs97_1><label for=boolIs97_1>Access 2000、2002、2003</label></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" 压 缩 " name="Submit"></td>
</tr>



</table>
</td>
</table>
</td></tr></table>
</form>

<%
end sub



sub statroom
set fso=server.createobject("scripting.filesystemobject")
upfiledir=server.mappath("./images/upfile")
set d=fso.getfolder(upfiledir)
upfilesize=d.size

toldir=server.mappath(".")
set d=fso.getfolder(toldir)
tolsize=d.size

totalBytes=upfacesize+upphotosize+upfilesize+tolsize



%>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>统计占用空间</td>
  </tr>
   <tr height=25>
    <td class=a3 width="29%">
    

上传附件占用空间</td>
    <td class=a3 width="69%">
    

<IMG height=10 src="images/bar/2.gif" width=<%=Int(upfilesize/totalBytes*100)%>%> <%totalsize(upfilesize)%></td></tr>



   <tr height=25>
    <td class=a3 width="29%">
    

BBSxp目录总共占用空间</td>
    <td class=a3 width="69%">
    

<IMG height=10 src="images/bar/4.gif" width=<%=Int(tolsize/totalBytes*100)%>%> <%totalsize(tolsize)%></td></tr></table>


<%
end sub


sub totalsize(size)
if size<1024 then
response.write size&" Bytes"
elseif size<1048576 then
response.write Int(size/1024)&" KB"
else
response.write Int(size/1024/1024)&" MB"
end if
end sub

sub files
thisdir=server.mappath("./images/upfile/"&dir&"")
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set fdir=fs.GetFolder(thisdir)

count=fdir.Files.count

%>

<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>

<a href="?menu=files&dir=<%=dir%>"><%=dir%>目录下</a>共有 <font color="FF0000"><b><%=count%></b></font> 个文件

<table border="0" width="100%"><tr><td valign=top>
<img src=images/folder.gif> <a href="?menu=files">..</a><br>

<%

set gFolder=fs.GetFolder(thisdir)

for each theFolder in gFolder.SubFolders
%>
<img src=images/folder.gif> <a href="?menu=files&dir=<%=theFolder.Name%>"><%=theFolder.Name%></a><br>
<%
next

%>

</td>
<td valign=top align="center">
<%

pagesetup=10 '设定每页的显示数量
If Count/pagesetup > (Count\pagesetup) then
TotalPage=(Count\pagesetup)+1
else TotalPage=(Count\pagesetup)
End If
if Request.QueryString("ToPage")<>"" then PageCount = cint(Request.QueryString("ToPage"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage


%>
<form method=post name=form action="?menu=delfiles">
<input type=hidden name=dir value=<%=dir%>>

<table cellspacing="1" cellpadding="2" border="0" width=88% class="a2">
<tr class=a1 ><td width=50 align="center">

<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
	<td width=208 align=center>名称</td>
  <td width=128 align=center>
大小</td><td width=140 align="center">类型</td><td align="center">
修改时间</td></tr>
<%
pagesize=20
page=request.querystring("page")
if page="" or not isnumeric(page) then
	page=1
else
	page=int(page)
end if
filenum=fdir.Files.count
pagenum=int(filenum/pagesize)
if filenum mod pagesize>0 then
	pagenum=pagenum+1
end if
if page> pagenum then
	page=1
end if
i=0


genre=Request("genre")

For each thing in fdir.Files


if instr(thing.name,genre) > 0 then

	i=i+1
	if i>(page-1)*pagesize and i<=page*pagesize then

response.write "<tr class=a3><td align=center><input type='CheckBox' value='"&thing.name&"' name=files></td><td align=center><a target=_blank href='images/upfile/"&dir&"/"&thing.Name&"'>"&thing.Name&"</a></td><td align=center>" &thing.size& " byte</td><td align=center>" & thing.type & "</td><td align=center>" & cstr(thing.datelastmodified) & "</td></tr>"

	elseif i>page*pagesize then
		exit for
	end if
end if

Next

%></table>


</div>


<table width="88%" cellpadding=2 cellspacing=2>
<tr><td>
<INPUT type=submit onclick="checkclick('您确定要删除您所选的文件?');" value=" 删 除 " name=Submit>

</td></form>

	<td align="center">
	
<form method=post name=form action="?menu=files">
<input type=hidden name=dir value=<%=dir%>>

搜索文件名包含
<input size="20" name="genre" value="<%=genre%>">
<INPUT type=submit value=" 搜 索 " name=Submit2></td>

	<td align="right">

<%
if page>1 then
	response.write "<a href=?menu=files&dir="&dir&"&genre="&genre&"&page=1>首页</a>&nbsp;&nbsp;<a href=?menu=files&dir="&dir&"&genre="&genre&"&page="& page-1 &">上一页</a>&nbsp;&nbsp;"
else
	response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
end if
if page<i/pagesize then
	response.write "<a href=?menu=files&dir="&dir&"&genre="&genre&"&page="& page+1 &">下一页</a>&nbsp;&nbsp;<a href=?menu=files&dir="&dir&"&genre="&genre&"&page="& pagenum &">尾页</a>"
else
	response.write "下一页&nbsp;&nbsp;尾页"
end if
%>
</td>

</tr></table>


</td>
	</tr>
</table>
<%

end sub

sub ReForum
%>
<form method=post name=form action="?menu=update">
<table  cellspacing="1" cellpadding="2" border="0" width="60%" class="a2">
	<tr class=a1>
		<td align="center" colspan="4" height="25">当前回帖数据表管理</td>
	</tr>
	<tr class=a3>
		<td align="center">表名</td>
		<td align="center">数据</td>
		<td align="center">默认</td>
		<td align="center">管理</td>
	</tr>
	<tr class=a3>
		<td align="center">ReForum</td>
		<td align="center"><%=conn.execute("Select count(id)from [ReForum]")(0)%></td>
		<td align="center"><input type="radio" <%if ReForumName=empty then%>CHECKED<%end if%> value="" name="TableName"></td>
		<td align="center"><font color="#C0C0C0">清理</font> <font color="#C0C0C0">删除</font></td>
	</tr>
	
<%
sql="select * from [TableList]"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
%>
	<tr class=a3>
		<td align="center">ReForum<%=rs("TableName")%></td>
		<td align="center"><%=conn.execute("Select count(id)from [ReForum"&rs("TableName")&"]")(0)%></td>
		<td align="center"><input type="radio" <%if ""&ReForumName&""=""&rs("TableName")&"" then%>CHECKED<%end if%> value="<%=rs("TableName")%>" name="TableName"></td>
		<td align="center"><a href="?menu=clean&tablename=<%=rs("TableName")%>">清理</a> <a href="?menu=del&tablename=<%=rs("TableName")%>">删除</a></td>
	</tr>
<%
rs.movenext
loop
Set Rs = Nothing
%>
	
	<tr class=a3>
		<td align="center" colspan="4">
		<input type="submit" value="提 交"></td>
	</tr>
	
	<tr class=a3>
		<td colspan="4">
		<i>注释：</i>上面数据表中选中的为当前论坛所使用来保存回帖数据的表，一般情况下每个表中的数据越少论坛帖子显示速度越快，当您下列单个数据表中的数据有超过几十万的帖子时不妨新添一个数据表来保存帖子数据，您会发现论坛速度快很多。 
		您也可以将当前所使用的数据表在数据表中切换，当前所使用的帖子数据表即当前论坛用户发贴时默认保存回帖的数据表.<br>
		<i>建议：</i>ACCESS 版本超过 200,000 条数据建立新表<br>
		　　　 MSSQL 版本超过 500,000 条数据建立新表</td>
	</tr>
</table></form>
<br>
<form method=post name=form action="?menu=CreatTable">
<table  cellspacing="1" cellpadding="2" border="0" width="60%" class="a2">

	<tr class=a1>
		<td align="center" colspan="2" height="25">
		添加数据表</td>
	</tr>
	<tr class=a3>
		<td align="center" width="70%">
		添加的表名:<input value="1" name="tablename" size="2" onkeyup=if(isNaN(this.value))this.value='' maxLength=1>&nbsp;&nbsp; 只需要填写数字( 1 - 9 )即可</td>
		<td align="center" width="30%">
		<input type="submit" value="提 交"></td>
	</tr>
	</table></form>
<%
end sub

htmlend
%>