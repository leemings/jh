<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim path
dim objFSO
dim uploadfolder
dim uploadfiles
dim upname
dim UpFolder
dim upfilename

dim admin_flag
admin_flag="72"
dim sfor(30,2)
if not master or instr(session("flag"),admin_flag)=0 then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_Error()
else
	if request("Submit")="清理记录" then
		call delall()
	elseif request("Submit")="清除未记录文件" then
		call delall1()
	else
		call main()
	end if
end if

sub main()
if request("path")<>"" then
path=request("path")
else 
path="UploadFile"
end if
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
            <tr >
              <td width="100%" valign=top colspan=2 class="forumrow"><p>
                    <b>注意</b>：<BR>①、本功能必须服务器支持FSO权限方能使用，FSO使用帮助请浏览微软网站。如果您服务器不支持FSO请手动管理。<BR>②、新版（DV6）之后的版本上传目录强制定义为UploadFile，只有该目录下文件可进行文件自动清理工作，新版之前的版本上传文件只能手动清除垃圾上传文件<br>③、自动清理文件：将对所有上传文件进行核实，如发现文件没有被相关帖子所使用，将执行自动清除命令
        </td>
    </tr>
                <tr> 
                  <td class="forumRowHighlight" colspan=2 height=25> 

<b>论坛上传文件夹</b>：<A HREF="?path=uploadimages">uploadimages</a>　｜　<A HREF="?path=UploadFile">UploadFile</a>　</font>
                  </td>
                </tr>
<tr > 
<form method="POST" action="?action=pathname">
                  <td width=20% class="forumrow">要查看的文件夹：</td><td width=80% class="forumrow">
<input type="text" name="path" size="35">&nbsp;<input type="submit" value="提交">
       （<font color=red>请填写正确的文件夹名或路径</font>）</td></form>
                </tr>
   <tr> 
                  <td colspan=2 height=25 class="forumRowHighlight"> 
                    <b>当前浏览　<%=path%>　目录的所有文件列表如下</b>
                  </td>
                </tr>
				</table><BR>
<%
sFor(0,0)="txt":sFor(0,1)="txt"
sFor(1,0)="chm":sFor(1,1)="chm"
sFor(2,0)="hlp":sFor(2,1)="chm"
sFor(3,0)="doc":sFor(3,1)="doc"
sFor(4,0)="pdf":sFor(4,1)="pdf"
sFor(5,0)="gif":sFor(5,1)="gif"
sFor(6,0)="jpg":sFor(6,1)="jpg"
sFor(7,0)="png":sFor(7,1)="png"
sFor(8,0)="bmp":sFor(8,1)="bmp"
sFor(9,0)="asp":sFor(9,1)="asp"
sFor(10,0)="jsp":sFor(10,1)="asp"
sFor(11,0)="js" :sFor(11,1)="asp"
sFor(12,0)="htm":sFor(12,1)="html"
sFor(13,0)="html":sFor(13,1)="html"
sFor(14,0)="shtml":sFor(14,1)="html"
sFor(15,0)="zip":sFor(15,1)="zip"
sFor(16,0)="rar":sFor(16,1)="rar"
sFor(17,0)="exe":sFor(17,1)="exe"
sFor(18,0)="avi":sFor(18,1)="avi"
sFor(19,0)="mpg":sFor(19,1)="mpg"
sFor(20,0)="ra" :sFor(20,1)="ra"
sFor(21,0)="ram":sFor(21,1)="ra"
sFor(22,0)="mid":sFor(22,1)="mid"
sFor(23,0)="wav":sFor(23,1)="wav"
sFor(24,0)="mp3":sFor(24,1)="mp3"
sFor(25,0)="asf":sFor(25,1)="asf"
sFor(26,0)="php":sFor(26,1)="aspx"
sFor(27,0)="php3":sFor(27,1)="aspx"
sFor(28,0)="aspx":sFor(28,1)="aspx"
sFor(29,0)="xls":sFor(29,1)="xls"
sFor(30,0)="mdb":sFor(30,1)="mdb"

dim pagesize, page, filenum, pagenum
pagesize=20
page=request.querystring("page")
if page="" or not isnumeric(page) then
	page=1
else
	page=int(page)
end if
%>

<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center style="table-layout:fixed;word-break:break-all">
    <tr  align=center>
<th width="50" height=25>类型</th>
<th width="*">文件地址</th>
<th width="100">大小</th>
<th width="130">最后访问</th>
<th width="130">上传日期</th>
<th width="30">管理</th>
</tr>
<%

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if request("filename")<>"" then
   if objFSO.fileExists(Server.MapPath(""&path&"\"&request("filename"))) then
        objFSO.DeleteFile(Server.MapPath(""&path&"\"&request("filename")))
    else
        response.write "未找到"&path&request("filename")
   end if
end if
on error resume next
Set uploadFolder=objFSO.GetFolder(Server.MapPath(""&path&"\"))
if err.number<>0 then
response.write "<tr><td colspan=6 class=forumrow>"&Err.Description&"</td></tr>"
response.end
end if
Set uploadFiles=uploadFolder.Files

filenum=uploadfiles.count
pagenum=int(filenum/pagesize)
if filenum mod pagesize>0 then
	pagenum=pagenum+1
end if
if page> pagenum then
	page=1
end if

i=0
For Each Upname In uploadFiles
	i=i+1
	if i>(page-1)*pagesize and i<=page*pagesize then
		upfilename=upname.name
                response.write "<tr><td align=center height=24 class=forumrow><img src=""images/files/"& procGetFormat(upname.name) &".gif"" border=0></td>"
		response.write "<td  class=forumrow><a href="""&path&"/"&upfilename&""" target=_blank>"&upfilename&"</a></td>"
		response.write "<td align=center  class=forumrow>"& upname.size &" B </td>"
		response.write "<td align=center  class=forumrow>"& upname.datelastaccessed &"</td>"
		response.write "<td align=center  class=forumrow>"& upname.datecreated &"</td>"
		response.write "<td align=center  class=forumrow><a href='?filename="&upname.name&"&path="&request("path")&"'>删除</a></td></tr>"
	elseif i>page*pagesize then
		exit for
	end if
next
set uploadFolder=nothing
set uploadFiles=nothing
%>
<tr>
<form method="POST" action="?path=<%=path%>" name=myfile>
<td colspan=6 align=center height=25 class="forumRowHighlight">
<%
if page>1 then
	response.write "<a href=?page=1&path="&request("path")&">首页</a>&nbsp;&nbsp;<a href=?page="& page-1 &"&path="&request("path")&">上一页</a>&nbsp;&nbsp;"
else
	response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
end if
if page<i/pagesize then
	response.write "<a href=?page="& page+1 &"&path="&request("path")&">下一页</a>&nbsp;&nbsp;<a href=?page="& pagenum &"&path="&request("path")&">尾页</a>"
else
	response.write "下一页&nbsp;&nbsp;尾页"
end if

response.write " 共 "&filenum&" 个文件  "&_
		"</td></tr>"
if path="UploadFile" then
%><tr><td colspan=6 align=center height=25 class="forumRowHighlight">
清理时同时直接删除文件
<input type=radio name=delfile value=1 >是&nbsp;
<input type=radio name=delfile value=2 checked>否
<input type="submit" name="Submit" value="清理记录"><input type="submit"  name="Submit" value="清除未记录文件" onclick="{if(confirm('您确定执行此操作吗?将删除所以未有记录的上传文件,并不能恢复。')){this.document.myfile.submit();return true;}return false;}">
</td></tr>
<%end if%>
</form>
</tr>
</table>
<%
	end sub

sub delall()
path="UploadFile"

dim F_ID,F_AnnounceID,F_boardid,F_filename
dim S_AnnounceID,s_Rootid
dim drs,delfile
dim delinfo
delfile=trim(request.form("delfile"))
if cint(delfile)=1 then
delinfo="已被删除！"
else
delinfo="未被删除！"
end if

i=0
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set rs=conn.execute("select F_ID,F_AnnounceID,F_BoardID,F_Filename from [DV_Upfile] where   F_Flag=0 order by F_ID desc ")
if rs.eof  then
	response.write "还未有"
else
	do while not rs.eof
	F_ID=rs(0)
	F_boardid=rs(2)
	if instr(rs(3),"/")=0 then '判断文件是否本论坛，若不是则采用表中的记录．
	F_filename="UploadFile/"&rs(3)
	else
	F_filename=rs(3)
	end if
	if rs(1)="" or isnull(rs(1)) then
		if delfile=1 then
		conn.execute("delete from DV_Upfile where F_ID="&F_ID&" ")
		end if
		if instr(rs(3),"/")=0 then
		if objFSO.fileExists(Server.MapPath(F_filename)) then
		if delfile=1 then
		objFSO.DeleteFile(Server.MapPath(F_filename))
		end if
		response.write "文件未写帖子,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
		else
		response.write "文件未写帖子,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> 已不存在！<br>"
		end if
		else
		response.write "外部文件<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
		end if
		i=i+1
	else
		if isnumeric(rs(1)) then
		S_AnnounceID=rs(1)
		else
		F_AnnounceID=split(rs(1),"|")
		s_Rootid=F_AnnounceID(0)
		S_AnnounceID=F_AnnounceID(1)
		end if
	
		'找出相应的帖子进行判断文件是否存在帖子内容
		set drs=conn.execute("select body from "&AllPostTable(0)&" where AnnounceID="&S_AnnounceID&" ")
		if drs.eof  then
			if delfile=1 then
			conn.execute("delete from DV_Upfile where F_ID="&F_ID&" ")
			end if
			if objFSO.fileExists(Server.MapPath(F_filename)) then
			if delfile=1 then
			objFSO.DeleteFile(Server.MapPath(F_filename))
			end if
			response.write "帖子未找到,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
			else
			response.write "帖子未找到,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> 已不存在！<br>"
			end if
			i=i+1
		else
			if instr(drs(0),"viewfile.asp?ID="&F_ID&"")=0 and instr(drs(0),F_filename)=0 then
			if delfile=1 then
			conn.execute("delete from DV_Upfile where F_ID="&F_ID&" ")
			end if
			if objFSO.fileExists(Server.MapPath(F_filename)) then
			if delfile=1 then
			objFSO.DeleteFile(Server.MapPath(F_filename))
			end if
			response.write "帖子内容不符,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"[<a href=""dispbbs.asp?Boardid="&F_boardid&"&ID="&s_Rootid&"&replyID="&S_AnnounceID&"&skin=1"" target=""_blank"" title=""浏览相关帖子""><font color=red>查看相关讨论</font></a> | <a href=myfile.asp?action=edit&editid="&F_ID&" target=""_blank"" title=""编辑文件""><font color=red>编辑</font></a>]<br>"
			else
			response.write "帖子内容不符,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> 已不存在！[<a href=""dispbbs.asp?Boardid="&F_boardid&"&ID="&s_Rootid&"&replyID="&S_AnnounceID&"&skin=1"" target=""_blank"" title=""浏览相关帖子""><font color=red>查看相关讨论</font></a> | <a href=myfile.asp?action=edit&editid="&F_ID&" target=""_blank"" title=""编辑文件""><font color=red>编辑</font></a>]<br>"
			end if
			i=i+1
			end if
		end if
		drs.close
	end if
rs.movenext
loop
end if
rs.close
set drs=nothing
set rs=nothing
set objFSO=nothing

response.write"共清理　"&i&"　个无用文件 ［<a href=?path="&path&" >返回</a>］"
end sub


sub delall1()
dim delfile,delinfo
delfile=checkStr(trim(request.form("delfile")))
if cint(delfile)=1 then
delinfo="已被删除！"
else
delinfo="未被删除！"
end if

path="UploadFile"
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set uploadFolder=objFSO.GetFolder(Server.MapPath(""&path&"\"))
Set uploadFiles=uploadFolder.Files
i=0
For Each Upname In uploadFiles
    upfilename=""&path&"/"&upname.name
     set rs=conn.execute("select top 1 F_ID from DV_Upfile where  F_Filename like '%"&upname.name&"%'   ")
     if rs.eof  then
     i=i+1
	if delfile=1 then
     	objFSO.DeleteFile(Server.MapPath(upfilename))
	end if
     response.write ""&upfilename&" "&delinfo&"<br>"
     end if
     rs.close
     set rs=nothing
next
response.write"共删除　"&i&"　个无用文件 ［<a href=?path="&path&" >返回</a>］"
set uploadFolder=nothing
set uploadFiles=nothing

end sub


sub delall2()
dim selectfile
dim bid,sid,did
dim delfile,delinfo
delfile=checkStr(trim(request.form("delfile")))
if cint(delfile)=1 then
delinfo="已被删除！"
else
delinfo="未被删除！"
end if

path=request("path")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set uploadFolder=objFSO.GetFolder(Server.MapPath(""&path&"\"))
Set uploadFiles=uploadFolder.Files
i=0
For Each Upname In uploadFiles
    upfilename=""&path&"/"&upname.name

if instr(upname.name,"_")>0 then
     selectfile=split(upname.name,"_")
     bid=selectfile(0)
     sid=selectfile(1)
if isnumeric(bid) and  isnumeric(sid) then
     set rs=conn.execute("select body from "&AllPostTable(0)&" where   AnnounceID="&sid&"  ")
     if rs.eof  then
     i=i+1
	if delfile=1 then
     	objFSO.DeleteFile(Server.MapPath(upfilename))
	end if
     	response.write ""&upfilename&" "&delinfo&"<br>"
      	else
           if instr(rs(0),upfilename)=0 then
        i=i+1
if delfile=1 then
     objFSO.DeleteFile(Server.MapPath(upfilename))
end if
     response.write ""&upfilename&" "&delinfo&"<br>"
            end if
     end if
     rs.close
     set rs=nothing
   end if
else
	i=i+1
if delfile=1 then
	objFSO.DeleteFile(Server.MapPath(upfilename))
end if
	response.write ""&upfilename&" 已被删除！<br>"
end if
next
response.write"共删除　"&i&"　个无用文件 ［<a href=?path="&path&" >返回</a>］"
set uploadFolder=nothing
set uploadFiles=nothing

end sub


function folder(path)
on error resume  next
       Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
          Set uploadFolder=objFSO.GetFolder(Server.MapPath(path))
		  if err.number<>"0" then
		  response.write Err.Description
		  response.end
		  end if
          For Each UpFolder In uploadFolder.SubFolders
            response.write "『<A HREF=?path="&path&"/"&upfolder.name&" >"&upfolder.name&"</a>』 | "
next
set uploadFolder=nothing
end function

function procGetFormat(sName)
 dim i,str
 procGetFormat=0
 if instrRev(sName,".")=0 then exit function
 str=lcase(mid(sName,instrRev(sName,".")+1))
 for i=0 to uBound(sFor,1)
  if str=sFor(i,0) then 
    procGetFormat=sFor(i,1)
    exit for
  end if
 next
end function
%>