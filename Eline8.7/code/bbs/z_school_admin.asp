<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: z_school_admin.asp
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
stats="ͬѧ¼����ҳ��"
call nav()
call head_var(0,0,boardtype,"z_school_class.asp?boardid="&boardid)
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�����ͬѧ¼��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> �����ͬѧ¼��������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
end if
if  boardmaster or  master or  superboardmaster then
founderr=false
else
founderr=true
Errmsg=Errmsg+"<br>"+"<li>ֻ�й���Ա���ܵ�¼��"
end if
if founderr then
call dvbbs_error()
else
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
response.write "<TABLE cellpadding=0 cellspacing=1 class=tableborder1 align=center>"
response.write "<tr><td height=20 class=TopLighNav1 align=center><a href=show.asp?filetype=1&boardid="&boardid&">�༶���</a> | <a href=z_school_classuser.asp?boardid="&boardid&">�༶��Ա</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>Ⱥ�����</a> | <a href=list.asp?boardid="&boardid&">�༶��̳</a> |  "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>�༶��ҳ</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">�ҵ�����</a> | <a href=announce.asp?boardid="&boardid&">��Ҫ����</a> | <a href=z_school_inclass.asp?boardid="&boardid&">����༶</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">�˳��༶</a> | <a href=z_school_admin.asp?boardid="&boardid&">�༶����</a></td></tr>"
response.write "<tr><td height=25 class=tablebody1 align=center>��ӭ"&htmlencode(membername)&"����ͬѧ¼����ҳ��<td></tr><tr><td height=20 class=TopLighNav1 align=center><a href=z_school_admin.asp?action=1&boardid="&boardid&">������Ϣ����</a> | <a href=z_school_admin.asp?action=2&boardid="&boardid&">�༶��Ա����</a> | <a href=z_school_admin.asp?boardid="&boardid&">�༶���淢��</a></td></tr><tr><td class=tablebody1>"
set rs=server.createobject("adodb.recordset")
Select Case request("action")
Case 1     '������Ϣ����
call editbminfo()
Case 2     '�༶��Ա����
call class1()
Case 3
call class2()
Case 4   'д����(1)
call savenews()
Case 11   '������Ϣ����(1)
call savebminfo()
Case 22    '�༶��Ա����(2)
call class2()
Case else    'д����
call news()
end Select
if founderr then call dvbbs_error()
response.write "</td></tr></TABLE>"
end if
call activeonline()
call footer()
'------------�༶��Ա����-----------------------------------
sub class1()
const zdes=20
dim page,pga,pgb
if request("page")="" or not isnumeric(request("page")) then
	page=1
else
	page=clng(request("page"))
end if
dim boarduser,userzz
userzz="<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style='width:100%'><form action='z_school_admin.asp?action=22&boardid="&boardid&"' method=post name=FORM>"
set rs=conn.execute("select boarduser from board where boardid="&boardid)
if rs(0)<>"" then
boarduser=split(replace(rs(0)," ",""),",")
else
	response.write"��ͬѧ¼��û��ͬѧ��"
	exit sub
end if
set rs=nothing
dim userinfo,a,pgc,pgd
pgd=ubound(boarduser)+1
pgc="��������"&pgd
a=1
userzz=userzz&"<tr height=25><th>����</th><th>�ա�����</th><th>E-MAIL</th><th>�硡����</th><th>�ء�����������ַ</th><th>�ǡ�����</th></tr>"
if pgd > zdes then
pga=zdes*(page-1)
pgb=zdes*page-1
pgc=pgc&"����ҳ��"
for i = 1 to (pgd+zdes-1)/zdes 
if i = page then
pgc=pgc&" [<font color=#FF0000>"&i&"</font>] "
else
pgc=pgc&" [<a href=z_school_admin.asp?action=2&boardid="&boardid&"&page="&i&">"&i&"</a>] "
end if
next
else
pga=0
pgb=pgd-1
end if
if pgb > ubound(boarduser) then
pgb=ubound(boarduser)
end if
dim ss,dd,ff,gg
for i = pga to pgb
if boarduser(i)<>"" then
set rs=conn.execute("select userid,useremail,userinfo from [User] where username='"&boarduser(i)&"'")
if rs(2)<>"" then
userinfo=split(rs(2),"|||")
ss=userinfo(0)
dd=userinfo(5)
ff=userinfo(13)
gg=userinfo(14)
else
ss=" "
dd=" "
ff=" "
gg=" "
end if
userzz=userzz&"<tr height=25><td class=tablebody"&a&" align=center><input type=checkbox name=boarduser value='"&boarduser(i)&"' ></td><td class=tablebody"&a&"><a href=DISPUSER.ASP?id="&rs(0)&" target=_blank>"&ss&"("&boarduser(i)&")</a></td><td class=tablebody"&a&">"&rs(1)&"</td><td class=tablebody"&a&">"&ff&"</td><td class=tablebody"&a&">"&gg&"</td><td class=tablebody"&a&">"&dd&"</td></tr>"
set rs=nothing
end if
if a=1 then a=2 else a=1
next
userzz=userzz&"<tr height=25><td class=tablebody1 colspan=5>"&pgc&"</td><td class=tablebody2 align=center><input type=Submit value='�� ��' name=Submit></td></tr></form></TABLE>"
response.write userzz
end sub
'------------�༶��Ա����(2)-----------------------------------
sub class2()
dim boarduser,username
username=replace(request("boarduser")," ","")
sql="select boarduser,txlpd from board where boardid="+Cstr(request("boardid"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>������������"
	founderr=true
	exit sub
elseif cint(left(rs("txlpd"),1))=0 then
	Errmsg=Errmsg+"<br>"+"<li>����ͬѧ¼���������"
	founderr=true
	exit sub
elseif isnull(rs(0)) or rs(0)="" then
	Errmsg=Errmsg+"<br>"+"<li>����ͬѧ¼������"
	founderr=true
	exit sub
else
	boarduser=rs(0)
end if
if instr(username,",") = 0 then
	username=" "&username&" "
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	boarduser=replace(boarduser,username," ")
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
else
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	dim dede
	username=split(username, ",")
	for each dede in username
		dede=" "&dede&" "
		boarduser=replace(boarduser,dede," ")
	next
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
end if
conn.execute("update board set boarduser='"&boarduser&"' where boardid="+Cstr(request("boardid")))
response.write "�����ɹ���"&request("boarduser")
end sub
'-----------------------------------------------
sub class3()



end sub
'------------------д����-----------------------------
sub news()
%>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:100%">
<form action="z_school_admin.asp?action=4" method=post name=FORM>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">�����ˣ� </div>
      </td>
      <td width="80%" class=tablebody1>
        <input type=text name=username size=36 value="<%=membername%>" disabled>
        <input type=hidden name=username value="<%=membername%>">
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">���⣺ </div>
      </td>
      <td width="80%" class=tablebody1>
        <input type=text name=title size=60><input type=hidden name="boardid" value="<%=CStr(BoardID)%>">
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">���ݣ� </div>
      </td>
      <td width="80%" class=tablebody1>
        <textarea cols=60 rows=6 name="content"></textarea>
      </td>
    </tr>
    <tr>
      <td width="100%" valign=top colspan="2" align=center class=tablebody2> 
        <input type=Submit value="�� ��" name=Submit>
        &nbsp; 
        <input type="reset" name="Clear" value="�� ��">
      </td>
    </tr>
  </table>
<%
end sub
'-------------------д����(1)----------------------------
sub savenews()
	dim username,title,content
	if request("boardid")<>"" or (not isInteger(request("boardid"))) then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br>"+"<li>�������˴���Ĳ�����"
		founderr=true
	end if
	if request("username")="" then
		Errmsg=Errmsg+"<br>"+"<li>�����뷢���ߡ�"
		founderr=true
	else
		username=checkstr(request("username"))
	end if
	if request("title")="" then
		Errmsg=Errmsg+"<br>"+"<li>���������ű��⡣"
		founderr=true
	else
		title=checkstr(request("title"))
	end if
	if request("content")="" then
		Errmsg=Errmsg+"<br>"+"<li>�������������ݡ�"
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
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
	end if
end sub
sub success()
%><br><br>
�ɹ������Ų���
<%
end sub

'������Ϣ����
sub editbminfo()
dim master_1

%> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="+CSTr(boardid)
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
   call dvbbs_error()
exit sub
   end if
master_1=split(rs("boardmaster"),"|")
if not master then
	if membername<>master_1(0)  then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������Աר�á�"
	call dvbbs_error()
	exit sub
	else
	Errmsg=Errmsg+"<br>"+"<li>��δ���޸����õ�Ȩ�ޡ�"
	call dvbbs_error()
	exit sub
	end if
end if
Forum_Setting = split(rs("Board_Setting"),",")
%>
             
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:100%">
<form action ="z_school_admin.asp?action=11&boardid=<%=boardid%>" method=post>
    <tr> 
      <th width="20%" height=22 class=tablebody2><b>�ֶ����ƣ�</b> </td>
      <th  > 
        <div align="center"><b>����ֵ��</b></div>
      </th>
    </th>
    <tr> 
      <td height=22 class=tablebody1  align="center">�༶���ƣ�</td>
      <td  class=tablebody1>
	  <input type="text" name="boardtype" size="30" value='<%=htmlencode(rs("boardtype"))%>'>
	  <input type='hidden' name=editid value='<%=boardid%>'></td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2  align="center">����˵����</td>
      <td  class=tablebody1> 
        <input type="text" name="readme" size="60" value='<%=htmlencode(rs("readme"))%>'>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody1  align="center">������Ա��</td>
      <td  class=tablebody1> 
        <input type="text" name="boardmaster" size="30" value='<%
if instr(rs("boardmaster"),"|") <> 0 then
response.write replace(rs("boardmaster"),master_1(0)&"|","")
end if
%>'>(�ั����Ա�������|�ָ����磺׷����|wxz)
      </td>
    </tr>
<tr> 
      <td height=22 class=tablebody1  align="center">�Ƿ񿪷ţ�</td>
      <td  class=tablebody1>
<input type=radio name="Board_Setting(2)" value="0" <%if cint(Forum_Setting(2))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(2)" value="1" <%if cint(Forum_Setting(2))=1 then%>checked<%end if%>>��&nbsp;(�������Ǳ����ͬѧҲ���Խ���!)
      </td>
    </tr><%
dim txlpd
txlpd=split(rs("txlpd"),"$")
if txlpd(1)="" then
txlpd(1)=0
end if
%>
<tr> 
      <td height=22 class=tablebody1  align="center">��ѧ��ݣ�</td>
      <td  class=tablebody1>
<select name="txlpd(1)">
<%for i = 1955 to 2003%>
<option value="<%=i%>" <%if i=cint(txlpd(1)) then%>selected<%end if%>><%=i%></option>
<%next%>
</select>
      </td>
    </tr>
<tr> 
      <td height=22 class=tablebody1  align="center">�༶��ҳ��</td>
      <td  class=tablebody1>http://<input type="text" name="txlpd(2)" size="30" value='<%=txlpd(2)%>'>
     </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2>&nbsp;</td>
      <td  class=tablebody2> 
        <input type="submit" name="Submit" value="�ύ">
      </td>
    </tr></form>
  </table>
<%
rs.close
end sub


'������Ϣ����(1)
sub savebminfo()
dim rname,zmzm,boarduser,zmrr,txlpd,txlpd1
set rs= server.CreateObject ("adodb.recordset")
sql = "select boardmaster,Board_Setting,boarduser,txlpd from board where boardid="+CSTr(request("boardid"))
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
   call dvbbs_error()
exit sub
end if
if not master then
	rname=split(rs(0),"|")
	if membername<>rname(0)  then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������Աר�á�"
	call dvbbs_error()
	exit sub
	else
	Errmsg=Errmsg+"<br>"+"<li>��δ���޸����õ�Ȩ�ޡ�"
	call dvbbs_error()
	exit sub
	end if
end if
txlpd=split(rs(3),"$")
zmrr=split(checkStr(rs(0)),"|")
boarduser=rs(2)
zmzm=rs(1)
set rs=nothing
if CSTr(Request("Board_Setting(2)"))=1 then
zmzm=left(zmzm,4)&"1"&mid(zmzm,6)
elseif CSTr(Request("Board_Setting(2)"))=0 then
zmzm=left(zmzm,4)&"0"&mid(zmzm,6)
end if
rname=split(checkStr(Request("boardmaster")),"|")
for i=0 to ubound(rname)
sql="select top 1 username from [user] where username='"&replace(rname(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>��̳û������û����������Ϊ������Ա"
	exit sub
	exit for
elseif instr(boarduser,","&rname(i))=0 and instr(boarduser,rname(i)&",")=0 then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>ͬѧ¼û������û����������Ϊ������Ա"
	exit sub
	exit for
end if
set rs=nothing
next
txlpd1=txlpd(0)&"$"&checkreal(request.Form("txlpd(1)"))&"$"&checkreal(request.Form("txlpd(2)"))&"$"&txlpd(3)&"$"&txlpd(4)&"$"&txlpd(5)&"$1"
rname = zmrr(0)&"|"&checkStr(Request("boardmaster"))
set rs=server.createobject("adodb.recordset")
sql = "select * from board where boardid="+Cstr(request("boardid"))
rs.open sql,conn,1,3
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
if Request("boardmaster")<>"" then
	rs("boardmaster") = rname
end if
	rs("readme") = checkStr(Request("readme"))
	rs("boardtype")=checkStr(Request("boardtype"))
	rs("Board_Setting")= zmzm
	rs("txlpd")= txlpd1
	rs.Update 
	rs.Close 
	response.write "<p>��̳�޸ĳɹ���"
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"$","��")
	checkreal=w
end if
end function
%>




