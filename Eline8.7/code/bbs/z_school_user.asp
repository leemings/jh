<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/z_school_char.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/birthday.asp"-->
<%
'=========================================================
' File: z_school_user.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'========================================================= 
dim suserid
if request("id")="" then
	suserid=userid
elseif isnumeric(request("id")) then
	suserid=request("id")
else
	suserid=session("userid")
end if
stats="��Ա����"
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)

if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�����ͬѧ¼��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> �����ͬѧ¼��������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
elseif cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>��ͬѧ¼Ϊ��֤ͬѧ¼����<a href=login.asp>��¼</a>��ȷ�����Ǳ���ͬѧ��"
		founderr=true
	else
		if chkschoollogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>�㲻�Ǳ���ͬѧ������<a href=z_school_inclass.asp?boardid="&boardid&">���뱾��</a>��"
		founderr=true
		end if
	end if
end if
if founderr then
call dvbbs_error()
else
call class1()
end if
call activeonline()
call footer()
sub class1()
dim boarduser,userzz
response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
response.write "<col width=20% ><col width=*><col width=40% >"
sql="select boarduser,txlpd from board where boardid="&boardid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
response.write "<tr><td height=20 class=TopLighNav1 align=center colspan=3><a href=show.asp?filetype=1&boardid="&boardid&">�༶���</a> | <a href=z_school_classuser.asp?boardid="&boardid&">�༶��Ա</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>Ⱥ�����</a> | <a href=list.asp?boardid="&boardid&">�༶��̳</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>�༶��ҳ</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">�ҵ�����</a> | <a href=announce.asp?boardid="&boardid&">��Ҫ����</a> | <a href=z_school_inclass.asp?boardid="&boardid&">����༶</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">�˳��༶</a> | <a href=z_school_admin.asp?boardid="&boardid&">�༶����</a></td></tr>"
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where userid="&suserid
rs.open sql,conn,1,1
if err.number<>0 then 
	ErrMsg=Errmsg+"<br>"+"<li>���ݿ����ʧ�ܣ�"&err.description
	founderr=true
	exit sub
end if
if rs.eof and rs.bof then
	ErrMsg=Errmsg+"<br>"+"<li>����ѯ�����ֲ�����"
	founderr=true
	exit sub
end if
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,useraddress,userinfo
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
		useraddress=userinfo(14)
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
		useraddress=""
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
	useraddress=""
end if
%> 
  <tr height=25> 
    <th colspan=2 align=left>��������</th>
    <td rowspan=9 align=center class=tablebody1 width=40% valign=top>
<%=iimg(rs("userphoto"),"<font color=gray>��</font>","<img src="&htmlencode(FilterJS(rs("userPhoto")))&">")%>
    </td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�� �ƣ�</td>
    <td class=tablebody2><%=htmlencode(rs("username"))%></td>
</tr>   
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>�� ��</td>
    <td class=tablebody1><%=iif(rs("sex"),"��","Ů")%> </td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�� ����</td>
    <td class=tablebody2>
<%=iimg(rs("birthday"),"<font color=gray>δ��</font>",rs("birthday"))%>
 </td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>�� ����</td>
    <td class=tablebody1>
<%=iimg(rs("birthday"),"<font color=gray>δ��</font>",astro(rs("birthday")))%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�ţ���죺</td>
    <td class=tablebody2>
<%=iimg(rs("useremail"),"<font color=gray>δ��</font>","<a href=mailto:"&htmlencode(rs("useremail"))&">"&htmlencode(rs("useremail"))&"</a>")%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>�� �ѣ�</td>
    <td class=tablebody1>
<%=iimg(rs("oicq"),"<font color=gray>δ��</font>",htmlencode(rs("oicq")))%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�ɣãѣ�</td>
    <td class=tablebody2>
<%=iimg(rs("icq"),"<font color=gray>δ��</font>",htmlencode(rs("icq")))%>
</td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>�ͣӣΣ�</td>
    <td class=tablebody1>
<%=iimg(rs("msn"),"<font color=gray>δ��</font>",htmlencode(rs("msn")))%>
 </td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�� ҳ��</td>
    <td class=tablebody2>
<%=iimg(rs("homepage"),"<font color=gray>δ��</font>","<a href="&htmlencode(rs("homepage"))&" target=_blank>"&htmlencode(rs("homepage"))&"</a> ")%>
</td><td class=tablebody1 align=center width=40% >
      <b><a href="javascript:openScript('messanger.asp?action=new&touser=<%=htmlencode(rs("username"))%>',500,400)">��������</a> | <a href="friendlist.asp?action=addF&myFriend=<%=HTMLEncode(rs("username"))%>" target=_blank>��Ϊ����</a></b></td>
  </tr>
<tr height=25> 
    <th colspan=2 align=left>
      �û���ϸ����</th>
    <td rowspan=14 class=tablebody1 width=40% valign=top>
<b>�ԣ�����</b>
<br>
<%=character%>
<br><br><br>
<b>���˼�飺</b><br>
<%=personal%>
<br>
</td>
  </tr>   
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>��ʵ������</td>
    <td class=tablebody1><%=realname%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�������ң�</td>
    <td class=tablebody2><%=country%> </td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>��ϵ�绰��</td>
    <td class=tablebody1><%=userphone%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>ͨ�ŵ�ַ��</td>
    <td class=tablebody2><%=useraddress%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>ʡ�����ݣ�</td>
    <td class=tablebody1><%=province%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>�ǡ����У�</td>
    <td class=tablebody2><%=city%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>������Ф��</td>
    <td class=tablebody1><%=shengxiao%> </td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>Ѫ�����ͣ�</td>
    <td class=tablebody2><%=blood%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>�š�������</td>
    <td class=tablebody1><%=belief%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>ְ����ҵ��</td>
    <td class=tablebody2><%=occupation%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>����״����</td>
    <td class=tablebody1><%=marital%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody2 width=20% align=right>���ѧ����</td>
    <td class=tablebody2><%=education%></td>
  </tr>
  <tr height=25> 
    <td class=tablebody1 width=20% align=right>��ҵԺУ��</td>
    <td class=tablebody1><%=college%></td>
  </tr>
<%
rs.close
set rs=nothing
call activeonline()
response.write "</TABLE>"
end sub
%>