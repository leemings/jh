<!-- #include file="conn.asp" -->
<%
Set Rs=Conn.Execute("clubconfig")
clubname=rs("clubname")
cluburl=rs("cluburl")
adminpassword=rs("adminpassword")
homeurl=rs("homeurl")
homename=rs("homename")
affichetitle=rs("affichetitle")
badwords=rs("badwords")
badip=rs("badip")
selectmail=rs("selectmail")
smtp=rs("smtp")
smtpmail=rs("smtpmail")
MailServerusername=rs("MailServerusername")
MailServerPassword=rs("MailServerPassword")
banner=rs("banner")
adcode=rs("adcode")
badlist=rs("badlist")
nowdate=rs("nowdate")
today=rs("today")
oldday=rs("oldday")
style=rs("style")
CacheName=rs("CacheName")
selectup=rs("selectup")
UpFileGenre=rs("UpFileGenre")
ReForumName=rs("ReForumName")
bbsxp_setup=split(rs("allclass"),"\")
floor=bbsxp_setup(0)
CloseRegUser=bbsxp_setup(1)
Reg10=bbsxp_setup(2)
RegOnlyMail=bbsxp_setup(3)
SendPassword=bbsxp_setup(4)
apply=bbsxp_setup(5)
Timeout=bbsxp_setup(6)
OnlineTime=bbsxp_setup(7)
MaxFace=bbsxp_setup(8)
MaxPhoto=bbsxp_setup(9)
MaxFile=bbsxp_setup(10)
searchcontent=bbsxp_setup(11)
MaxSearch=bbsxp_setup(12)
ActivationForum=bbsxp_setup(13)
ActivationUser=bbsxp_setup(14)
SortShowForum=bbsxp_setup(15)
PostTime=bbsxp_setup(16)
ListStyle=bbsxp_setup(17)
StopOutPost=bbsxp_setup(18)
set rs=nothing

if Request.Cookies("username") <> empty then
sql="select * from [user] where username='"&HTMLEncode(Request.Cookies("username"))&"'"
Set Rs=Conn.Execute(SQL)
if rs.eof then Response.Cookies("username")=""
if Request.Cookies("userpass") <> rs("userpass") then Response.Cookies("username")=""
membercode=rs("membercode")
userface=""&rs("userface")&""
newmessage=rs("newmessage")
userlife=rs("userlife")
set rs=nothing
end if

if badip<>empty then
filtrate=split(badip,"|")
for i = 0 to ubound(filtrate)
if instr("|"&Request.ServerVariables("REMOTE_ADDR")&"","|"&filtrate(i)&"") > 0 then response.redirect "inc/badip.htm"
next
end if

if Request.Cookies("skins")=empty then Response.Cookies("skins")=style


dim ForumTreeList
ii=0
startime=timer()
Set rs = Server.CreateObject("ADODB.Recordset")
Server.ScriptTimeout=Timeout '���ýű���ʱʱ�� ��λ:��

response.write "<html><head><meta http-equiv=Content-Type content=text/html;charset=gb2312></head><link href=images/skins/"&Request.Cookies("skins")&"/bbs.css rel=stylesheet><script src=inc/BBSxp.js></script><script src=inc/ybb.js></script><script src=images/skins/"&Request.Cookies("skins")&"/bbs.js></script>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function HTMLEncode(fString)
fString=replace(fString,";","&#59;")
fString=replace(fString,"<","&lt;")
fString=replace(fString,">","&gt;")
fString=replace(fString,"\","&#92;")
fString=replace(fString,"--","&#45;&#45;")
fString=replace(fString,"'","&#39;")
fString=replace(fString,CHR(34),"&quot;")
fString=replace(fString,vbCrlf,"<br>")
HTMLEncode=fString
end function
''''''''''�滻ģ��START''''''''''''
Function ReplaceText(fString,patrn,replStr)
Set regEx = New RegExp  ' ����������ʽ��
regEx.Pattern = patrn   ' ����ģʽ��
regEx.IgnoreCase = True ' �����Ƿ����ִ�Сд��
regEx.Global = True     ' ����ȫ�ֿ����ԡ� 
ReplaceText = regEx.Replace(fString, replStr) ' ���滻��
Set reg=nothing
End Function
''''''''''�滻ģ��END''''''''''''
function ContentEncode(fString)
fString=replace(fString,vbCrlf, "")
fString=replace(fString,"\","&#92;")
fString=replace(fString,"'","&#39;")
'fString=ReplaceText(fString,"<(.[^>]*)(&#|cookie|window.|Document.|javascript:|js:|vbs:|about:|file:|on(blur|click|change|Exit|error|focus|finish|key|load|mouse))", "&lt;$1$2$3")
fString=ReplaceText(fString,"<(.[^>]*)(Document.|onerror|onload|onmouseover)", "&lt;$1$2")
fString=ReplaceText(fString,"<(\/|)(iframe|SCRIPT|form|style|div|object|TEXTAREA)", "&lt;$1$2")
ContentEncode=fString
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub top
if Request.Cookies("username")<>empty then
''''''''''''''''''��һ����'''''''''''''''''''''''''''''
if Request.Cookies("onlinetime")=empty then
conn.execute("update [user] set degree=degree+1,landtime='"&now()&"' where username='"&Request.Cookies("username")&"'")
Response.Cookies("onlinetime")=now()
Response.Cookies("addmin")=0
''''''''���ͣ��10���ӣ�����ֵ: +1 ����ֵ��+10''''''''''''''''''''
elseif DateDiff("s",Request.Cookies("onlinetime"),Now())>600 then

if userlife > 90 then
conn.execute("update [user] set userlife=100,landtime='"&now()&"' where username='"&Request.Cookies("username")&"'")
else
conn.execute("update [user] set userlife=userlife+10,landtime='"&now()&"' where username='"&Request.Cookies("username")&"'")
end if

Response.Cookies("onlinetime")=now()
Response.Cookies("addmin")=Request.Cookies("addmin")+10
end if
''''''''''''''''''''''''''''
end if
%>
<body topmargin=0>
<script>bbsxptop()</script>
<table cellspacing=1 cellpadding=1 width=97% align=center border=0 class=a2><tr class=a1><td><table cellspacing=0 cellpadding=4 width=100% border=0><tr class=a1><td id=TableTitleLink>&gt;&gt;��ӭ���� <%
if Request.Cookies("username")=empty then
%>  <a href="login.asp">���ȵ�¼</a> |  <a href="register.asp">
	ע��</a> |  <a href="RecoverPasswd.asp">
	��������</a> |  <a href="online.asp">
	�������</a> | <a href="search.asp?forumid=<%=Request("forumid")%>">
	����</a> | <a href="help.asp">
	����</a> </td></tr></table></td></tr></table><br>
<%else%>
<%=Request.Cookies("username")%>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=cookies.asp?menu=online>����</a></div><div class=menuitems><a href=cookies.asp?menu=eremite>����</a></div><div class=menuitems><a href=login.asp?menu=out>����</a></div>')" style=cursor:default>
<%
if Request.Cookies("eremite")="1" then
response.write "����"
else
response.write "����"
end if
%></A>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=ShowBBS.asp?menu=5>�ҵ�����</a></div><div class=menuitems><a href=calendar.asp>������־</a></div><div class=menuitems><a href=blog.asp>�ҵ���־</a></div><div class=menuitems><a href=profile.asp>�ҵ�����</a></div>')" style=cursor:default>���˷���</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=usercp.asp>���������ҳ</a></div><div class=menuitems><a href=EditProfile.asp>���������޸�</a></div><div class=menuitems><a href=EditProfile.asp?menu=pass>�û������޸�</a></div><div class=menuitems><a href=EditProfile.asp?menu=option>�༭��̳ѡ��</a></div><div class=menuitems><a href=message.asp>�û����ŷ���</a></div><div class=menuitems><a href=friend.asp>�༭�����б�</a></div>')" style=cursor:default>�������</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=online.asp>�������</a></div><div class=menuitems><a href=online.asp?menu=cutline>����ͼ��</a></div><div class=menuitems><a href=online.asp?menu=sex>�Ա�ͼ��</a></div><div class=menuitems><a href=online.asp?menu=todaypage>����ͼ��</a></div><div class=menuitems><a href=online.asp?menu=board>����ͼ��</a></div><div class=menuitems><a href=online.asp?menu=tolrestore>����ͼ��</a></div><div class=menuitems><a href=usertop.asp>��Ա�б�</a></div><div class=menuitems><a href=adminlist.asp>�����Ŷ�</a></div>')" style=cursor:default>��̳״̬</a>
 | <a href=search.asp>����</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=favorites.asp?menu=topic>�����ղؼ�</a></div><div class=menuitems><a href=favorites.asp?menu=forum>��̳�ղؼ�</a></div><div class=menuitems><a href=favorites.asp>��վ�ղؼ�</a></div>')" style=cursor:default>�ղ�</a>

<%
menu(0)

response.write " | <a href=help.asp>����</a>"

select case membercode
case "5"
response.write " | <a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=admin.asp target=_top>��¼����</a></div><div class=menuitems><a href=recycle.asp>�� �� վ</a></div>')"&Chr(34)&" style=cursor:default>����</a>"
case "4"
response.write " | <a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=recycle.asp>�� �� վ</a></div>')"&Chr(34)&" style=cursor:default>��������</a>"
end select


%></td><td align="right">�Ѿ�ͣ�� <b><%=DateDiff("n",Request.Cookies("onlinetime"),Now())+Request.Cookies("addmin")%></b> ����&nbsp; </td></tr></table></td></tr></table>
<%
'''''����ж�ѶϢ''''''''''''''''''''''
if newmessage>0 then
%><table width="97%" align="center"><tr><td width="100%" align="right"><a href=# onclick="javascript:open('friend.asp?menu=look&<%=now()%>','','width=320,height=170')"><img src=images/newmail.gif border=0><font color="990000">����<%=newmessage%>����ѶϢ����ע�����</font></a> </td></tr></table>
<script src=inc/Messenger.js></script>
<script>javascript:getMsg("&nbsp;BBSxp Messenger","&nbsp;&nbsp;<a style=cursor:hand href=# onclick=\"javascript:open('friend.asp?menu=look&<%=now()%>','','width=320,height=170')\">����<%=newmessage%>����ѶϢ����ע����գ�</a>");</script>
<%
else
response.write "<br>"
end if
'''''''''''''
end if

end sub
''''''''''''''''''''''''''''''
sub htmlend
%><title><%=clubname%> - Powered By BBSxp</title><p>
<table cellspacing=0 cellpadding=0 width=97% align=center><tr><td align=middle>
<%=adcode%><br>Powered by <font color=ffffff> <a target=_blank href=http://www.bbsxp.com/download.htm><font color=000000>BBSxp <%=ver%></font></a></font>/<a  href=Licence.asp><font color=000000>Licence</font></a>&nbsp;&copy; 1998-2005<br>
Script Execution Time:<%=fix((timer()-startime)*1000)%>ms
</td></tr></table>
<script>bbsxpbottom()</script>
</body></html>
<%
responseend
end sub
''''''''''''''''''''''''''''''''
sub succeed(message)
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� ������Ϣ</td>
</tr>
</table><br><SCRIPT>valigntop()</SCRIPT>
<table cellspacing="0" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr><td height="106">
<table cellspacing="1" cellpadding="6" width="100%" border="0"><tr><td width="100%" height="20" align="middle" class="a1">
<b>������ʾ��Ϣ</b></td></tr><tr bgcolor="ffffff"><td valign="top" align="left" height="122"><b>&nbsp;</b><table cellspacing="0" cellpadding="5" width="100%" border="0"><tr><td width="83%" valign="top"><b><span id="yu">3</span><a href="javascript:countDown"></a>���Ӻ�ϵͳ���Զ�����...</b><ul><%=message%></ul></td><td width="17%"><img height="97" src="images/succ.gif" width="95"></td></tr></table></td></tr></table></td></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
</font><script>function countDown(secs){yu.innerText=secs;if(--secs>0)setTimeout("countDown("+secs+")",1000);}countDown(3);</script>
<%
htmlend
end sub
sub error(message)
response.redirect "error.asp?message="&message&""
end sub
sub error2(message)
response.redirect "error.asp?message2="&message&""
end sub
''''''''''''''''''''''''''''''''
sub log(message)
conn.Execute("insert into log (username,ip,windows,remark) values ('"&Request.Cookies("username")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.Servervariables("HTTP_USER_AGENT")&"','"&message&"')")
end sub
''''''''''''''''''''''''''''''''
sub responseend
set rs=nothing
set rs1=nothing
conn.close
set conn=nothing
Response.End
end sub
''''''''''''''''''''''''''''''''

sub DetectPost
if StopOutPost = 1 then
if instr(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&"") = 0 then error("<li>��Դ����")
end if
end sub


sub ForumTree(selec)
sql="Select * From bbsconfig where id="&selec&""
Set Rs1=Conn.Execute(sql)
if not rs1.eof then
ForumTreeList="<span id=temp"&selec&"><a onmouseover=loadtree('"&selec&"') href=ShowForum.asp?forumid="&Rs1("id")&">"&Rs1("bbsname")&"</a></span> �� "&ForumTreeList&""
ForumTree Rs1("followid")
end if
Set Rs1 = Nothing
end sub



sub menu(selec)
sql="Select * From menu where followid="&selec&" order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
if rs1("followid")=0 then 
%> | <a onmouseover="showmenu(event,'<%menu(rs1("id"))%>')" style=cursor:default><%=rs1("name")%></a><%
else
response.write "<div class=menuitems><a href="&rs1("url")&">"&rs1("name")&"</a></div>"
end if
rs1.movenext
loop
Set Rs1 = Nothing
end sub

sub ClubTree
sql="Select * From bbsconfig where followid=0 and hide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
ClubTreeList=ClubTreeList&"<div class=menuitems><a href=ShowForum.asp?forumid="&Rs1("id")&">"&Rs1("bbsname")&"</a></div>"
rs1.movenext
loop
Set Rs1 = Nothing
response.write "<a onmouseover="&Chr(34)&"showmenu(event,'"&ClubTreeList&"')"&Chr(34)&" href=Default.asp>"&clubname&"</a>" 
end sub

sub BBSList(selec)
sql="Select * From bbsconfig where followid="&selec&" and hide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
Response.write "<option value="&rs1("ID")&">"&string(ii,"��")&""&rs1("bbsname")&"</option>"
ii=ii+1
BBSList rs1("id")
ii=ii-1
rs1.movenext
loop
Set Rs1 = Nothing
end sub

if Application(CacheName&"CountForum")=empty then Application(CacheName&"CountForum")=conn.execute("Select count(id) from [forum]")(0)
if Application(CacheName&"CountReforum")=empty then Application(CacheName&"CountReforum")=conn.execute("Select sum(replies) from [forum]")(0)
if Application(CacheName&"CountReforum")<1 then Application(CacheName&"CountReforum")=0
if Application(CacheName&"CountUser")=empty then Application(CacheName&"CountUser")=conn.execute("Select count(id) from [user]")(0)


%>