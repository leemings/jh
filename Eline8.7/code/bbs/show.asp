<!--#include FILE="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
stats="��̳չ��"
dim TopicCount
dim Pcount,endpage,star,page_count
dim toback,tonext
dim rsearch
dim maxpage
dim tab
tab=request.cookies("tab")

dim bbsurl
bbsurl=getservepath(request.ServerVariables("server_name")&request.ServerVariables("URL"))

call nav()
if request("star")="" or not isnumeric(request("star"))  then
	star=1
        toback=1
        tonext=2
else
	star=clng(request("star"))
        toback=star-1
        tonext=star+1
end if
if star <=1 then 
	star=1
        toback=1
        tonext=2
end if

if tab="" or not  isnumeric(tab)  then
tab=4
end if

if request("tab")="" or not isnumeric(request("tab"))  then
	tab=tab
else
	tab=clng(request("tab"))
	Response.Cookies("tab").Expires=Date+365
	response.cookies("tab")=tab
end if

maxpage=clng(tab*3)	'ÿҳ��ʾ�ļ��ĸ���

if request("username")="" or  request("filetype")="" or request("boardid")=""  then
rsearch=""
end if

if request("boardid")<>"" and isnumeric(request("boardid")) then
	if  clng(request("boardid"))<>0 then
		rsearch=rsearch&"and  F_BoardID="&clng(request("boardid"))&"  "
	end if
end if

if request("filetype")<>"" and isnumeric(request("filetype"))  then
rsearch=rsearch&"and  F_Type="&cint(request("filetype"))&"  "
end if

if not master and not superboardmaster then
	rsearch=rsearch&"and  F_Username='"&membername&"'  "
else
	if request("username")<>""  then
		rsearch=rsearch&"and  F_Username='"&checkStr(request("username"))&"'  "
	end if
end if

if Cint(GroupSetting(49))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���������̳չ����Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
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
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()
%>
<%
sub main()
%> 
<table cellpadding=3 cellspacing=1 align=center class=tableborder1  >
    <tr>
      <th  height=25 align=center colspan="2"><%=stats%></th>
    </tr>
    <tr>
      <td width="180"  valign="top" class=tablebody2>
<%
'===================STAR TopInfo==================
response.write"<table cellpadding=2 cellspacing=1  class=tableborder2  style=""width:100%;word-break:break-all;"">"&_
						"<tr>"&_
						"<th  height=25 align=center >TOP 10 չ����������</th>"&_
						"</tr>"
call top(10,"F_ViewNum")
response.write"</table>"

response.write"<br><table cellpadding=2 cellspacing=1 class=tableborder2  style=""width:100%;word-break:break-all;"">"&_
						"<tr>"&_
						"<th  height=25 align=center >TOP 10 չ����������</th>"&_
						"</tr>"
call top(10,"F_ID")
response.write"</table>"
'===================end TopInfo==================

response.write"</td><td width=*  valign=""top"" class=tablebody1>"

call searchlist()
'===================STAR FILEINFO================
response.write"<table cellpadding=1 cellspacing=1 border=0 width=100% style=""word-break:break-all;"">"&_
						"<tr>"
'ȡ������
dim frs
set frs=conn.execute("select count(F_id) from DV_Upfile where F_Filename<>'' "&rsearch&" ")
TopicCount=frs(0)
frs.close
set frs=nothing

'�������
dim t
dim F_ID,F_AnnounceID,F_BoardID,F_Filename,F_FileType,F_Type,F_Readme,F_DownNum,F_ViewNum
dim F_typename,showfile,golist
dim edit
edit=false
t=0
set rs=server.createobject("ADODB.recordset")
rs.open "select  *  from [DV_Upfile] where F_Filename<>'' "&rsearch&" order by F_ID desc",conn,1,1
if rs.eof and rs.bof then
response.write"<tr >"&_
						"<td height=26 class=tablebody2 align=center><font color="&Forum_body(8)&">���鿴���ļ��ѱ�ɾ�����ֹչʾ��</font></td>"&_
						"</tr >"
else
if TopicCount mod Cint(maxpage)=0 then
     		Pcount= TopicCount \ Cint(maxpage)
  	else
     		Pcount= TopicCount \ Cint(maxpage)+1
  	end if
	RS.MoveFirst
	if star > Pcount then star = Pcount
	if star < 1 then star = 1
	rs.PageSize = maxpage
	rs.AbsolutePage=star
	page_count=0
do while not rs.eof and page_count < Cint(maxpage)

F_ID=rs("F_ID")
F_AnnounceID=trim(rs("F_AnnounceID"))
F_BoardID=rs("F_BoardID")
F_Filename=trim(rs("F_Filename"))
F_FileType=trim(rs("F_FileType"))
F_Type=rs("F_Type")
F_Readme=trim(rs("F_Readme"))
F_DownNum=rs("F_DownNum")
F_ViewNum=rs("F_ViewNum")

if instr(F_Filename,"/")=0 then '�ж��ļ��Ƿ���̳������������ñ��еļ�¼��
	F_Filename=bbsurl&"UploadFile/"&F_Filename
end if

if not isnull(F_AnnounceID) and F_AnnounceID<>"" then
		F_AnnounceID=split(F_AnnounceID,"|")
		golist="<a href=""dispbbs.asp?Boardid="&F_BoardID&"&ID="&F_AnnounceID(0)&"&replyID="&F_AnnounceID(1)&"&skin=1"" target=""_blank"" title=""����������"">�鿴�������...</a>"
else
		golist="<font color=gray>��û���������</font>"
end if

	select case F_Type
	case 1
	F_typename="ͼƬ��"
	showfile="<img src="""&htmlencode(FilterJS(F_Filename))&"""  style='border: 1 solid #000000' width=90 height=90>"
	case 2
	F_typename="FLASH��"
	showfile="<img src=Images/files/info.gif border=0><br><img src=""images/files/"&F_FileType&".gif"" border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""�������"">FLASH����</a>"
	case 3
	F_typename="���ּ�"
	showfile="<img src=Images/files/info.gif border=0><br><img src=""images/files/"&F_FileType&".gif"" border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""�������"">��������</a>"
	case 4
	F_typename="��Ӱ��"
	showfile="<img src=Images/files/info.gif border=0><br><img src=""images/files/"&F_FileType&".gif"" border=0><a href="&htmlencode(FilterJS(F_Filename))&"  target=""_blank"" title=""�������"">��Ӱ����</a>"
	case else
	F_typename="�ļ���"
	showfile="<img src=Images/files/info.gif border=0><br><img src=""images/files/"&F_FileType&".gif"" border=0><a href=viewfile.asp?ID="&F_ID&"  target=""_blank"" title=""�������"">�������</a>"
	end select


if GroupSetting(48)=1 then
if  master or superboardmaster or boardmaster  then
edit=true
elseif rs("F_Username")=membername then
edit=true
else
edit=false
end if
end if


response.write "<td width=""25%"" height=""200"" align=""center"" class=tablebody1>"&_
			"<table cellpadding=2 cellspacing=1 class=tableborder1 style=""width:100%;height=100%;word-break:break-all;"" >"&_
			"<tr><td width=""100%"" height=""130"" align=""center"" valign=top class=tablebody1>"
'ִ���ļ�����ͼ��
if founduser then
	response.write"<table border=0 width=100%  cellspacing=0 cellpadding=0>"&_
		      "<tr><td width=* align=left>"
	if edit=true then
	response.write"<a href=myfile.asp?action=edit&editid="&F_ID&" title=""�༭���ļ�"" ><img src="&Forum_info(7)&"editfile.gif  border=0 width=""10"" height=""10""></a>"
	response.write"&nbsp;<a href=myfile.asp?action=fdel&delid="&F_ID&" title=""ɾ�����ļ�"" ><img src="&Forum_info(7)&"delete.gif  border=0 width=""10"" height=""10""></a>"
	end if
	response.write"</td><td width=20 align=right><a href=fileshow.asp?action=send&id="&F_ID&" ><img src="&Forum_info(7)&"newmail.gif border=0 title=""���͸�����""></a></td></tr></table>"
end if
response.write"<br><a href=""fileshow.asp?boardid="&F_BoardID&"&id="&F_ID&"""  target=""_blank"" title='�������ļ�'>"&showfile&"</a><br>"
response.write"</td></tr><tr><td class=tablebody2 >"&_
			"����: <a href=dispuser.asp?name="&trim(rs("F_Username"))&" >"&htmlencode(rs("F_Username"))&"</a>"&_
			"<br>����: <a href=show.asp?filetype="&F_Type&" >"&F_typename&" ["
if F_Type=0 then
response.write"<font color="&Forum_body(8)&">"&F_DownNum&"</font>"
else
response.write"<font color="&Forum_body(8)&">"&F_ViewNum&"</font>"
end if
response.write"]</a><br>˵��: "
if rs("F_Readme")<>"" or not isnull(rs("F_Readme")) then
if len(rs("F_Readme"))>26 then
response.write ""&left(htmlencode(replace(rs("F_Readme"),chr(10)," ")),26)&"..."
else
response.write Server.htmlencode(rs("F_Readme"))
end if
else
response.write "<font color=gray>δ�м�¼</a>"
end if
response.write "<br>"&golist
response.write "</td></tr>"&_
						"</table></td>"
	if t=tab-1 and t<tab then response.write "</tr><tr>"
	if t>tab-1 then 
		t=1
	else
		t=t+1
	end if
	page_count = page_count + 1
rs.movenext
loop
end if
rs.close
set rs=nothing
response.write "</tr></table>"
'===================END FILEINFO================
call list()

response.write "</td></tr>"&_
						"<tr><th  height=25 align=center colspan=""2"">"&stats&"</th>"&_
						"</tr></table>"
end sub

'��ҳ����
SUB LIST()
response.write "<BR><table cellpadding=0 cellspacing=3 border=0 width=100% align=center ><form method=post action=""?boardid="&request("boardid")&"&filetype="&request("filetype")&"&username="&request("username")&"&tab="&tab&""" ><tr><td valign=middle nowrap>չ������<b>"&TopicCount&"</b>����Ʒ��ÿҳ"&maxpage&"���ļ�����<b><font color=red>"&Pcount&"</font></b>ҳ"
response.write "</td><td align=center></td><td align=right>"
call DispPageNum(clng(star),PCount,"""?star=","&boardid="&request("boardid")&"&filetype="&request("filetype")&"&username="&request("username")&"&tab="&tab&"""")
response.write " ת��:<input type=text name=star size=3 maxlength=10  value='"& star &"'><input type=submit value=Go  id=button1 name=button1 >"     
response.write "</td></tr></form></table>"
end sub


'����������ת��
sub searchlist()
%>
 <table cellpadding=6 cellspacing=1 align=center width="100%">
<form method=POST action="?" >
    <tr>
      <td align=center width="*" >
<%if master or superboardmaster then
	if request("username")<>"" then
		response.write "���<font color="&Forum_body(8)&" >"&htmlencode(request("username"))&"</font>�ĸ���չ��"
	end if
end if%>
</td>
 <td  width="200" align="right">
<%if master or superboardmaster then%>
<input type="text" name="username" size="20" value="<%=htmlencode(request("username"))%>"><input type="submit" name="Submit"  value="�û���ѯ">
<%end if%>
</td>               
      <td  width="150" align="right">
	  <select name=filetype onchange='javascript:submit()'>
	  <option value="" >���ļ��������</option>
<%
dim iupload
iupload="�ļ���|ͼƬ��|FLASH��|���ּ�|��Ӱ��"
iupload=split(iupload,"|")
for i=0 to ubound(iupload)
response.write "<option value="&i&"  "
if request("filetype")<>"" and isnumeric(request("filetype")) then
if i=cint(request("filetype")) then
response.write "selected"
end if
end if
response.write ">"&iupload(i)&"</option>"
next
%>
</td>
      <td width="120"  align="right"><select name=boardid onchange='javascript:submit()'>
	<option selected>ѡ�����̳...</option>
	<%  	
dim ii
ii=0
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
response.write "<option value="&rs(0)&" "
if boardid=rs(0) then
response.write "selected"
end if
response.write ">"
select case rs(2)
case 0
response.write "��"
case 1
response.write "&nbsp;&nbsp;��"
case 2
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 3
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 4
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 5
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 6
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 7
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 8
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
case 9
response.write "&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��"
end select
response.write rs(1)&"</option>"
ii=0
rs.MoveNext
loop
set rs=nothing
response.write"</select></td>"
 %>
 <td width=20>
  <select name=tab onchange='javascript:submit()'>
	  <option value="" >����</option>
	  <option value="2" >2��</option>
	  <option value="3" >3��</option>
	  <option value="4" >4��</option>
	  </select>
	  </td>
    </tr>
</form>
  </table>
  <%
end sub

'top ���е���
sub top(num,str)
dim fileinfo
set rs=conn.execute("select top "&num&" F_ID,F_BoardID,F_Username,F_Filename,F_Readme from DV_Upfile order by "&str&" desc,F_addTime desc")
do while not rs.EOF
if rs(4)<>""  then
fileinfo=htmlencode(rs(4))
else
fileinfo=htmlencode(FilterJS(rs(3)))
end if
response.write "<tr><td valign=top class=tablebody1>"
response.write "<font face=Webdings>4</font><a href=""fileshow.asp?boardid="&rs(1)&"&id="&cint(rs(0))&"""  target=""_blank"" title='�������ļ�'>"
if len(fileinfo)>12 then
response.write left(replace(fileinfo,chr(10)," "),12)
else
response.write fileinfo
end if
response.write"</a>"
'response.write"&nbsp;&nbsp;<a href=dispuser.asp?name="&trim(rs(2))&"  >"&htmlencode(rs(2))&"</a>"
response.write"</td></tr>"
rs.MoveNext
loop
set rs=nothing
end sub

function getservepath(str)
dim tmpstr
tmpstr=split(str,"/")
getservepath="http://"&replace(str, tmpstr(ubound(tmpstr)), "")
end function%>
