<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="roomconfig.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if sfkf=0 then
	response.write "���ɷ��䴴�������ѹرա�"
	response.end
end if
if session("aqjh_inthechat")<>"1" then
	Response.Write "<script language=JavaScript>{alert('������ڽ��������Һ�ſ��Դ������ž���������');window.close();}</script>"
	Response.End
end if
if nowinroom<>0 then
	Response.Write "<script language=JavaScript>{alert('������ڴ���ʱ�ſ��Դ������ž���������');window.close();}</script>"
	Response.End
end if
function jc(z)
	if not isnumeric(z) or (z<>0 and z<>1) then
		jc=true
	else
		jc=false
	end if
end function
h=clng(request.form("bh"))	'����h
f=clng(request.form("fzxz"))	'����f
g=clng(request.form("sjxz"))	'�¼�g
i=clng(request.form("kp"))	'��Ƭi
j=clng(request.form("db"))	'�Ĳ�j
c=clng(request.form("xz"))	'����c
if jc(bh) or jc(fzxz) or jc(sjxz) or jc(kp) or jc(db) or jc(xz) then
	Response.Write "<script language=JavaScript>{alert('���ؾ��棬�ύ���ݴ��󣡣�');window.close();}</script>"
	Response.End 
end if

'�̶�ֵ
b=200	'��߿��Խ��������

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���,����,���,��Ա��,grade from �û� where ����='"&aqjh_name&"'",conn,2,2
mymp=rs("����")
mysf=rs("���")
myjinbi=rs("���")
myyinliang=rs("����")
myjinka=rs("��Ա��")
rs.close
if mysf<>"����" or aqjh_grade<5 or mymp="����" or mymp="����" or mymp="����" or mymp="�ٸ�" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<font size=2>�㲻�����Ż����������������ҡ��������ٸ��ģ�</font>"
	response.end
end if
e="����=''"&mymp&"''"
if xz=1 then
	d="ֻ��"&mymp&"���ɵ��ӷ��ɽ���"
	k=mymp&"�صأ��Ǳ�����Ա�ô��ߣ�ɱ�޺գ�����"
else
	d="�κ��˶����Խ���"
	k=mymp&"���������㽻�������ѣ���ӭ���ǹ��١�"
end if
if myjinbi<jinbi or myyinliang<yinliang or myjinka<jinka then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�������ɾ�������Ҫ���� ������"&yinliang&"����ң�"&jinbi&"����Ա�𿨣�"&jinka&"�����Ǯ���������ܴ�����"
	response.end
end if
rs.open "select count(����) as ���� from �û� where ����='"&mymp&"' and �ȼ�>="&dzdj,conn
zs=rs("����")
rs.close
if zs<dzrs then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�����ɵ������ȼ���<font color=red>"&dzdj&"</font>�����ϵ�ֻ��<font color=red><b>"&zs&"</b></font>�ˣ�����<font color=red><b>"&dzrs&"</b></font>�������Դ������䡣"
	response.end
end if
rs.open "select * from r where a='"&mymp&"'",conn,2,2
if not(rs.eof or rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��һ������ֻ���Դ���һ�����䣬�������Ѿ������������ˣ�');window.close();}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����=����-"&yinliang&",���=���-"&jinbi&",��Ա��=��Ա��-"&jinka&" where ����='"&aqjh_name&"'"
conn.execute "insert into r(a,b,c,d,e,f,g,h,i,j,k) values ('"&mymp&"',"&b&","&c&",'"&d&"','"&e&"',"&f&","&g&","&h&","&i&","&j&",'"&k&"')"
rs.open "select * from r order by id",conn,2,2
application.lock()
roomming=""
do while Not rs.Eof
	rooming=rooming&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&"|"&rs("l")&"|"&rs("m")&";"
	if rs("a")=mymp then
		id=rs("id")
	end if
	rs.MoveNext
loop
application("aqjh_room")=rooming
aqjh_roominfo=split(Application("aqjh_room"),";")
dim nameindex(0)
for roomsn=0 to ubound(aqjh_roominfo)-1
	fenroom=split(aqjh_roominfo(roomsn),"|")
	if fenroom(0)=mymp then
		application("aqjh_chatroomname"&roomsn)=fenroom(0)
		Application("aqjh_onlinelist"&roomsn)=nameindex
		Application("aqjh_useronlinename"&roomsn)=""
	end if
next
application.unlock()
gg="��ϲ<font color=blue>"&mymp&"</font>����<font color=red><b>"& aqjh_name &"</b></font><font color=blue><b>�����ĵ���Ŭ����չ���ţ�"&dzdj&"�����ϵĵ������ﵽ"&dzrs&"���ϣ�</b></font>վ���ؽ����䴴���˱��ŵ����ɾ�������<font color=red><b>"&mymp&"</b></font>��ϣ��"&aqjh_name&"�Ժ����Ŭ��������һ�����ӻԻ͵�ҵ����<br>�ؽ������Һ󼴿ɽ���"&mymp&"���������"
ltssay="<table bgcolor=FFFF00 width=80% align=center border=1 cellspacing=0 cellpadding='0'><tr><td bgcolor=00FF00 align=center>���ɾ��������������</td><tr><td><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
says=ltssay		'��������
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../hc3w_Admin/setup.css">
<title>���ɾ����������</title></head>
<body text="#FFFFFF" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
  <font color="#000000" size="2">
  <p>���������ɾ���������</p>
  </font></div>
<table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="75%">
  <tr> 
    <td width="67"> 
      <div align="center"><font color="#FFFFFF" size=2>����ɣ�</font></div>
    </td>
    <td width="197"><font size="2"> <%=id%></font></td>
    <td width="81"> 
      <div align="center"><font size="2" color="#FFFFFF">�� �� ��</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2"><%=mymp%></font></td>
  </tr>
  <tr> 
    <td width="67"> 
      <div align="center"><font color="#FFFFFF" size="2">��������</font></div>
    </td>
    <td width="197">
      <%if f=0 then%>
      ����
      <%else%>
      ��ֹ
      <%end if%>
    </td>
    <td width="81" nowrap> 
      <div align="center"><font color="#FFFFFF" size="2">�¼�����</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2">
      <%if g=0 then%>
      ����
      <%else%>
      ��ֹ
      <%end if%>
      </font></td>
  </tr>
  <tr> 
    <td width="67" nowrap> 
      <div align="center"><font color="#FFFFFF" size="2">��&nbsp;&nbsp;��</font></div>
    </td>
    <td width="197"><font color="#FFFFFF" size="2">
      <%if h=0 then%>
      ����
      <%else%>
      ��ֹ
      <%end if%>
      </font></td>
    <td width="81"> 
      <div align="center"><font color="#FFFFFF" size="2">��&nbsp;&nbsp;Ƭ</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2">
      <%if h=0 then%>
      ����
      <%else%>
      ��ֹ
      <%end if%>
      </font></td>
  </tr>
  <tr> 
    <td width="67"> 
      <div align="center"><font size="2" color="#FFFFFF">��&nbsp;&nbsp;��</font></div>
    </td>
    <td width="197"><font color="#FFFFFF" size="2">
      <%if j=0 then%>
      ����
      <%else%>
      ��ֹ
      <%end if%>
      </font></td>
    <td width="81"> 
      <div align="center"><font color="#FFFFFF" size="2">��&nbsp;&nbsp;��</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2"><%=d%></font></td>
  </tr>
  <tr> 
    <td colspan="4">
      <div align="center"> <font size="2"> <br>
        ��ϲ��ɹ����������ѵ����ɾ�������<font color=red><b><%=mymp%></b></font> ��������ֻҪ�ؽ���һ�������ҾͿ��Կ������ɵķ���</font><br>
        <br>
        <input  onClick="javascript:window.close()" value="�� ��" type=button name="close">
      </div>
    </td>
  </tr>
</table>