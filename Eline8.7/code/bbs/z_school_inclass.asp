<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/z_school_char.asp" -->
<%
'=========================================================
' File: z_school_inclass.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
dim z_membername
if request("username")<>"" then
	z_membername=checkstr(request("username"))
else
	z_membername=membername
end if
if request("action")="out" then
stats="�˳�ͬѧ¼"
elseif request("action")="xg" or request("action")="xg1" then
stats="�޸�����"
else
stats="����ͬѧ¼"
end if
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�����ͬѧ¼��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
call dvbbs_error()
else
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
response.write "<tr><td height=20 class=TopLighNav1 align=center colspan=2><a href=show.asp?filetype=1&boardid="&boardid&">�༶���</a> | <a href=z_school_classuser.asp?boardid="&boardid&">�༶��Ա</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>Ⱥ�����</a> | <a href=list.asp?boardid="&boardid&">�༶��̳</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>�༶��ҳ</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">�ҵ�����</a> | <a href=announce.asp?boardid="&boardid&">��Ҫ����</a> | <a href=z_school_inclass.asp?boardid="&boardid&">����༶</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">�˳��༶</a> | <a href=z_school_admin.asp?boardid="&boardid&">�༶����</a></tD></tr>"

if request("action")="out" then
response.write outschool(boardid,z_membername)
elseif request("action")="updat" or request("action")="xg1" then
call update()
else
call inclasszl()
end if


if founderr then call dvbbs_error()
response.write "</TABLE>"
end if
call activeonline()
call footer()



Rem �û�����ͬѧ¼
function inschool(boardid,username)
dim boarduser,inschool1
inschool1=true
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	inschool="������������"
	inschool1=false
elseif cint(right(rs(1),1))=0 then
	inschool="����ͬѧ¼���������"
	inschool1=false
elseif isnull(rs(0)) or rs(0)="" then
	boarduser=username
else
	boarduser=split(rs(0), ",")
	for i = 0 to ubound(boarduser)
		if trim(boarduser(i))=trim(username) then
			inschool="�����벻Ҫ�ظ������ͬѧ¼��"
			inschool1=false
		exit for
		end if
	next
	boarduser=username&","&rs(0)
end if
if inschool1 then
	conn.execute("update board set boarduser='"&boarduser&"' where boardid="&boardid)
	inschool="���ѳɹ������ͬѧ¼��"
end if
rs.close
set rs=nothing
inschool="<TR><TH height=25 colspan=2>ϵ ͳ �� ʾ</TH></tR><tr><td class=tablebody1 height=50 align=center colspan=2>"&inschool&"</td></tr>"
end function


Rem �û��˳�ͬѧ¼
function outschool(boardid,username)
dim boarduser,inschool1
inschool1=0
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	outschool="������������"
	inschool1=1
elseif cint(right(rs(1),1))=0 then
	outschool="����ͬѧ¼���������"
	inschool1=1
elseif isnull(rs(0)) or rs(0)="" then
	outschool="����ͬѧ¼������"
	inschool1=1
else
	boarduser=rs(0)
end if
if instr(username,",") <> 0 and inschool1=0 then
inschool1=2
end if
if inschool1=0 then
	username=" "&username&" "
	boarduser=replace(boarduser," ","")
	boarduser=","&boarduser&","
	boarduser=replace(boarduser,","," ")
	boarduser=replace(boarduser,username," ")
	boarduser=trim(boarduser)
	boarduser=replace(boarduser," ",",")
	conn.execute("update board set boarduser='"&boarduser&"' where boardid="&boardid)
	outschool="���ѳɹ��˳���ͬѧ¼��"
elseif inschool1=2 then
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
	conn.execute("update board set boarduser='"&boarduser&"' where boardid="&boardid)
	outschool="���ѳɹ��˳���ͬѧ¼��"
end if
outschool="<TR><TH height=25 colspan=2>ϵ ͳ �� ʾ</TH></tR><tr><td class=tablebody1 height=50 align=center colspan=2>"&outschool&"</td></tr>"
end function

sub inclasszl()
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>���û��������ڡ�"
	founderr=true
	exit sub
end if%>
<TR><TH height=25 colspan=2>�� д �� ��</TH></tR>
<form action="z_school_inclass.asp?action=<%if request("action")="xg" then%>xg1<%else%>updat<%end if%>&boardid=<%=boardid%>" method=POST name="theForm">
<%dim usersetting,setuserinfo,setusertrue
if rs("usersetting")<>"" then
	usersetting=split(rs("usersetting"),"|||")
	if ubound(usersetting)=1 then
	if isnumeric(usersetting(0)) then setuserinfo=usersetting(0) else setuserinfo=1
	if isnumeric(usersetting(1)) then setusertrue=usersetting(1) else setusertrue=0
	else
	setuserinfo=1
	setusertrue=0
	end if
else
	setuserinfo=1
	setusertrue=0
end if%>
<tr>    
<td width=40% class=tablebody1><B>�Ƿ񿪷����Ļ�������</B>��<BR>���ź���˿��Կ��������Ա�Email��QQ����Ϣ</td>
<td width=60% class=tablebody1>&nbsp;<input type=radio name=setuserinfo value=1  <%if setuserinfo=1 then%>checked<%end if%>>&nbsp;����&nbsp;&nbsp;<input type=radio name=setuserinfo value=0  <%if setuserinfo=0 then%>checked<%end if%>>&nbsp;������</td>
</tr>
<tr>    
<td width=40% class=tablebody1><B>�Ƿ񿪷�������ʵ����</B>��<BR>���ź���˿��Կ���������ʵ��������ϵ��ʽ����Ϣ</font></td>
<td width=60% class=tablebody1>&nbsp;<input type=radio name=setusertrue value=1 <%if setusertrue=1 then%>checked<%end if%>>&nbsp;����&nbsp;&nbsp;<input type=radio name=setusertrue value=0 <%if setusertrue=0 then%>checked<%end if%>>&nbsp;������</td>   
</tr>
<%dim userinfo
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,address
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		address=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		address=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	address=""
end if%>
<tr>
<th height=25 align=left valign=middle colspan=2>&nbsp;������ʵ��Ϣ���������ݽ�����д��</th>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>��ʵ������</b><input type=text name=realname size=18 value="<%=realname%>">&nbsp;<font color=red>(����)</font></td>
<td height=71 align=left valign=top  class=tablebody1 rowspan=14 width=60% >
<table width=100% border=0 cellspacing=0 cellpadding=5>
<tr>
<td class=tablebody1><b>�ԡ��� &nbsp; </b><br><%
	dim KidneyType,theKidney
	KidneyType="�����Ը�,������,��������,���ɵ�Ƥ,��������,���ÿɰ�,����ͨͨ,������,������,�ĵ�����,��������,�ƽ�����,��Ȥ��Ĭ,˼�뿪��,������ȡ,С�Ľ���,�����ѻ�,������ֱ,����ʧ��,�ó�����,��������,�����ɹ�,���û�ʧ,�����쿪,����Ƹ�,��������,��������,հǰ�˺�,ѭ�浸��,��������,���Կ���,���Թ���,��������,׷��̼�,���Ų��,�ƻ����,̰С����,����˼Ǩ,�������,ˮ���ﻨ,��ɫ����,��С����,��������,�¸�����,������ѧ,ʵ������,��ʵʵ��,��ʵ�ͽ�,Բ������,Ƣ������,����˹��,��ʵ̹��,��������,������"
	theKidney=split(KidneyType, ",")
	for i = 0 to ubound(theKidney)	
		response.write "<input type=""checkbox"" name=""character"" value="""&trim(theKidney(i))&""" "
		if instr(character,trim(theKidney(i)))>0 then '����д����Ը�
			response.write "checked" 
		end if
		response.write ">"&trim(theKidney(i))&" "
		if ((i+1) mod 5)=0 then response.write "<br>"  'ÿ����ʾ�����Ը���л���
	next   
%></td>
</tr>
<tr>
<td class=tablebody1><b>���˼�飺 &nbsp;</b><br><textarea name=personal rows=5 cols=90% ><%=personal%></textarea></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>�������ң�</b><b><input type=text name=country size=18 value="<%=country%>"></b></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>��ϵ�绰��</b><b><input type=text name=userphone size=18 value="<%=userphone%>"></b>&nbsp;<font color=red>(����)</font></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>ͨ�ŵ�ַ��</b><b><input type=text name=address size=18 value="<%=address%>"></b>&nbsp;<font color=red>(����)</font></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>ʡ�����ݣ�</b><input type=text name=province size=18 value="<%=province%>"></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>�ǡ����У�</b><input type=text name=city size=18 value="<%=city%>"></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>������Ф��</b><select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=ţ <%if shengxiao="ţ" then%>selected<%end if%>>ţ</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
</select>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>Ѫ�����ͣ�</b><select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=���� <%if blood="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>�š�������</b><select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=��� <%if belief="���" then%>selected<%end if%>>���</option>
<option value=���� <%if belief="����" then%>selected<%end if%>>����</option>
<option value=������ <%if belief="������" then%>selected<%end if%>>������</option>
<option value=������ <%if belief="������" then%>selected<%end if%>>������</option>
<option value=�ؽ� <%if belief="�ؽ�" then%>selected<%end if%>>�ؽ�</option>
<option value=�������� <%if belief="��������" then%>selected<%end if%>>��������</option>
<option value=���������� <%if belief="����������" then%>selected<%end if%>>����������</option>
<option value=���� <%if belief="����" then%>selected<%end if%>>����</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>ְ����ҵ��</b><select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="�ƻ�/����" <%if occupation="�ƻ�/����" then%>selected<%end if%>>�ƻ�/����</option>
<option value=����ʦ <%if occupation="����ʦ" then%>selected<%end if%>>����ʦ</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
<option value=����������ҵ <%if occupation="����������ҵ" then%>selected<%end if%>>����������ҵ</option>
<option value=��ͥ���� <%if occupation="��ͥ����" then%>selected<%end if%>>��ͥ����</option>
<option value="����/��ѵ" <%if occupation="����/��ѵ" then%>selected<%end if%>>����/��ѵ</option>
<option value="�ͻ�����/֧��" <%if occupation="�ͻ�����/֧��" then%>selected<%end if%>>�ͻ�����/֧��</option>
<option value="������/�ֹ�����" <%if occupation="������/�ֹ�����" then%>selected<%end if%>>������/�ֹ�����</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
<option value=��ְҵ <%if occupation="��ְҵ" then%>selected<%end if%>>��ְҵ</option>
<option value="����/�г�/���" <%if occupation="����/�г�/���" then%>selected<%end if%>>����/�г�/���</option>
<option value=ѧ�� <%if occupation="ѧ��" then%>selected<%end if%>>ѧ��</option>
<option value=�о��Ϳ��� <%if occupation="�о��Ϳ���" then%>selected<%end if%>>�о��Ϳ���</option>
<option value="һ�����/�ල" <%if occupation="һ�����/�ල" then%>selected<%end if%>>һ�����/�ල</option>
<option value="����/����" <%if occupation="����/����" then%>selected<%end if%>>����/����</option>
<option value="ִ�й�/�߼�����" <%if occupation="ִ�й�/�߼�����" then%>selected<%end if%>>ִ�й�/�߼�����</option>
<option value="����/����/����" <%if occupation="����/����/����" then%>selected<%end if%>>����/����/����</option>
<option value=רҵ��Ա <%if occupation="רҵ��Ա" then%>selected<%end if%>>רҵ��Ա</option>
<option value="�Թ�/ҵ��" <%if occupation="�Թ�/ҵ��" then%>selected<%end if%>>�Թ�/ҵ��</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>����״����</b><select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=δ�� <%if marital="δ��" then%>selected<%end if%>>δ��</option>
<option value=�ѻ� <%if marital="�ѻ�" then%>selected<%end if%>>�ѻ�</option>
<option value=���� <%if marital="����" then%>selected<%end if%>>����</option>
<option value=ɥż <%if marital="ɥż" then%>selected<%end if%>>ɥż</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>���ѧ����</b><select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=Сѧ <%if education="Сѧ" then%>selected<%end if%>>Сѧ</option>
<option value=���� <%if education="����" then%>selected<%end if%>>����</option>
<option value=���� <%if education="����" then%>selected<%end if%>>����</option>
<option value=��ѧ <%if education="��ѧ" then%>selected<%end if%>>��ѧ</option>
<option value=˶ʿ <%if education="˶ʿ" then%>selected<%end if%>>˶ʿ</option>
<option value=��ʿ <%if education="��ʿ" then%>selected<%end if%>>��ʿ</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>&nbsp;<b>��ҵԺУ��</b><input type=text name=college size=18 value="<%=college%>"></td>
</tr>
<tr height=25>
<td valign=top width=40% class=tablebody1>&nbsp;<B>OICQ���룺</B><input type=TEXT name="OICQ" size=18 value="<%if rs("oicq")<>"" then%><%=rs("oicq")%><%end if%>" maxlength=20></td>   
</tr>
<tr align="center"> 
<td colspan="2" width="100%" class=tablebody2><input type=Submit value="�� ��" name="Submit"> &nbsp; <input type="reset" name="Submit2" value="�� ��"></td>
</tr>
</form>
<%end sub


sub update()
if request("address")="" then
	errmsg=errmsg+"<br>"+"<li>ͨ�ŵ�ַ��ͬѧ¼�Ǳ���ġ�"
	founderr=true
	exit sub
end if
if trim(request("oicq"))<>"" then
	if not isnumeric(request("oicq")) or len(request("oicq"))>12 then
		errmsg=errmsg+"<br>"+"<li>Oicq����ֻ����4-12λ���֣���ͬѧ¼�Ǳ���ġ�"
		founderr=true
		exit sub
	end if
end if
dim realname
realname=trim(request("realname"))
realname=replace(realname," ","")
if realname="" or strLength(realname)<4 or strLength(realname)>8 then
	errmsg=errmsg+"<br>"+"<li>��ʵ������������"
	founderr=true
exit sub
elseif not isChinese(realname) then
	errmsg=errmsg+"<br>"+"<li>��ʵ����ӦΪ���֡�"
	founderr=true
	exit sub
end if
if request.form("userphone")="" then
	errmsg=errmsg+"<br>"+"<li>��ϵ�绰��ͬѧ¼�Ǳ���ġ�"
	founderr=true
	exit sub
end if
dim userinfo,usersetting
userinfo=checkreal(realname) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue")
if not isnull(request.form("oicq")) and request.form("oicq")<>"" then
	conn.execute("update [User] set oicq='"&request.form("oicq")&"' where userid="&userid)
end if
conn.execute("update [User] set userinfo='"&userinfo&"' where userid="&userid)
conn.execute("update [User] set usersetting='"&usersetting&"' where userid="&userid)
if request("action")="xg1" then
response.write "<TR><TH height=25 colspan=2>ϵ ͳ �� ʾ</TH></tR><tr><td class=tablebody1 height=50 align=center colspan=2>�޸ĳɹ���</td></tr>" 
else
response.write inschool(boardid,z_membername)
end if
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
end if
end function
%>
