<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: z_school_class.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body
stats="ͬѧ¼"
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�����ͬѧ¼��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> �����ͬѧ¼��������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
end if
if founderr then
call dvbbs_error()
else
call class1()
end if
call activeonline()
call footer()
sub class1()
set rs=conn.execute("SELECT BoardMaster,boarduser,txlpd FROM board where boardid="&boardid)
dim txlpd,classu,classu1,classu2
txlpd=split(rs(2),"$")
if rs(1)<>"" then
classu=split(rs(1), ",")
classu1=ubound(classu)+1
classu2=classu(0)
else
classu=1
classu1=0
classu2=""
end if
response.write "<TABLE cellpadding=0 cellspacing=1 class=tableborder1 align=center>"
response.write "<tr height=20><td class=TopLighNav1 align=center colspan=6><a href=show.asp?filetype=1&boardid="&boardid&">�༶���</a> | <a href=z_school_classuser.asp?boardid="&boardid&">�༶��Ա</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>Ⱥ�����</a> | <a href=list.asp?boardid="&boardid&">�༶��̳</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>�༶��ҳ</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">�ҵ�����</a> | <a href=announce.asp?boardid="&boardid&">��Ҫ����</a> | <a href=z_school_inclass.asp?boardid="&boardid&">����༶</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">�˳��༶</a> | <a href=z_school_admin.asp?boardid="&boardid&">�༶����</a></td></tr>"
response.write "<tr height=25><th colspan=2>�� �� �� Ϣ</th><th colspan=4>�� �� ͨ ��</th></tr>"
response.write "<tr><td class=tablebody1 colspan=2><table width=100% border=0 cellspacing=0 cellpadding=0><tr height=60><td width=50% valign=top>"&boardtype&"<br>���������� "&classu1&"<br>����ͬѧ��"&classu2&"<br>�� �� Ա��"&replace(rs(0),"|",",")&"</td><td width=50% valign=top>��ѧ��ݣ�"&txlpd(1)&"<br>�༶��ҳ��<a href=http://"&txlpd(2)&">����</a><br>�༶��ʼ�ˣ�"&txlpd(3)&"<br>�������ڣ�"&txlpd(0)&"</td></tr></table></td><td class=tablebody1 colspan=4  valign=top>"
set rs=nothing
set rs=conn.execute("select top 1 title,addtime,content from bbsnews where boardid="&boardid&"  order by id desc")
if rs.bof and rs.eof then
response.write "<CENTER>��ӭ������"&boardtype&"!</CENTER>"
else
response.write rs(0)&"("&rs(1)&")<br>"&rs(2)
end if
set rs=nothing
response.write "</td></tr>"
response.write "<tr><td class=tablebody1 colspan=5 height=20>"&_
	"<table width=100% >"&_
	"<tr>"&_
	"<td valign=middle height=20 width=50> <a href=AllPaper.asp?boardid="&boardid&" title=����鿴����̳����С�ֱ�><b>�㲥��</b></a> </td><td width=*> "&_
	"<marquee scrolldelay=150 scrollamount=4 onmouseout=""if (document.all!=null){this.start()}"" onmouseover=""if (document.all!=null){this.stop()}"">"
set rs=conn.execute("SELECT TOP 5 s_id,s_username,s_title FROM smallpaper where datediff('d',s_addtime,Now())<=1 and s_boardid="&boardid&" ORDER BY s_addtime desc")
do while not rs.eof
	response.write "��<font color="&Forum_pic(5)&">"&htmlencode(rs(1))&"</font>˵��<a href=javascript:openScript('viewpaper.asp?id="&rs(0)&"&boardid="&boardid&"',500,400)>"&htmlencode(rs(2))&"</a>"
rs.movenext
loop
set rs=nothing

response.write "</marquee><td align=right width=240><a href=elist.asp?boardid="& boardid &" title=�鿴���澫��><font color="&Forum_body(8)&"><B>����</B></font></a> | <a href=boardstat.asp?reaction=online&boardid="&boardid&" title=�鿴����������ϸ���>����</a> | <a href=bbseven.asp?boardid="&boardid&" title=�鿴�����¼�>�¼�</a> | <a href=BoardPermission.asp?boardid="&boardid&" title=�鿴�����û���Ȩ��>Ȩ��</a> | <a href=admin_boardset.asp?boardid="&boardid&">����</a>"
if isaudit=1 then
response.write " | <a href=admin_topiclist.asp?boardid="&boardid&">���</a>"
end if
response.write "</td></tr></table>"
%>
</td></tr>
<TR valign=middle>
<Th height=25 width=32>״̬</Th>
<Th width=*>�� ��</Th>
<Th width=80>�� ��</Th>
<Th width=64>�ظ�/����</Th>
<Th width=195>������ | �ظ���</Th>
</TR>
<%
sql="select top 10  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression from Topic where boardid="&boardid&" and locktopic<2 ORDER BY TopicID desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   response.write"<tr><td colspan=5>��û��ͬѧ���ԡ�</td></tr>"
else
       while (not rs.eof)
if not isnull(rs("lastpost")) then
	LastPost=split(rs("lastpost"),"$")
	if ubound(LastPost)=6 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=rs("dateandtime")
	body=htmlencode(LastPost(3))
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser=htmlencode(LastPost(0))
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=LastPost(5)
	LastRootid=rs("topicid")
	end if
else
	Lastuser="------"
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
end if

response.write"<TR><TD align=middle class=tablebody2 width=32>"
if rs(0)=1 then
response.write"<img src="&Forum_info(7)&Forum_statePic(2)&" alt=""������������"">"
else
response.write"<img src="&Forum_info(7)&Forum_statePic(1)&" >"
end if
response.write"</TD><TD class=tablebody1 width=*><a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid")&"&skin=1' target=_blank>"
response.write"<img src="&Forum_info(8)&""
if instr(rs("Expression"),Forum_Setting(44))>0 then
response.write rs("Expression")
else
response.write "face1.gif"
end if
response.write  " border=0 alt=""���´������������""></a>"&_ 
				"<a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid") &"&skin=1'>"
if len(rs("title"))>26 then
response.write ""&left(htmlencode(replace(rs("title"),chr(10)," ")),26)&"..."
else
response.write htmlencode(rs("title"))
end if
response.write "</a></TD>"&_
				"<TD align=middle  class=tablebody2  width=80><a href=""dispuser.asp?id="&rs("postuserid")&""">"&htmlencode(rs("postusername"))&"</a></TD>" 

response.write "<TD class=tablebody1 width=64 align=center>"
if rs("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&rs("votetotal")&"</b></font>  Ʊ"
else
	response.write rs("child") &"/"& rs("hits")
end if
response.write "</TD>" 
%>
<TD  class=tablebody2 width=195>&nbsp;<a href='dispbbs.asp?boardid=<%=rs("boardid")%>&id=<%=LastRootID%>&<%=LastID%>' >
<%=FormatDateTime(LastTime,2)%>&nbsp;<%=FormatDateTime(LastTime,4)%></a>
&nbsp;<font color="<%=Forum_body(8)%>">|</font>&nbsp;
<a href="dispuser.asp?id=<%=LastUserID%>" target=_blank><%=htmlencode(LastUser)%></a>
</TD>
</TR> 
<%
          rs.movenext
        wend
end if
response.write "</table>"
end sub
%>