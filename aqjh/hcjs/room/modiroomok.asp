<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="roomconfig.asp"-->
<!--#include file="../../chk.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
ChkPost()
if sfkf=0 then
	response.write "���ɷ��䴴�������ѹرա�"
	response.end
end if
if session("aqjh_inthechat")<>"1" then
	Response.Write "<script language=JavaScript>{alert('������ڽ��������Һ�ſ����޸ı��ž����������ã���');window.close();}</script>"
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���,grade from �û� where ����='"&aqjh_name&"'",conn,2,2
mymp=rs("����")
mysf=rs("���")
rs.close
if mysf<>"����" or aqjh_grade<5 or mymp="����" or mymp="����" or mymp="����" or mymp="�ٸ�" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�㲻�����ţ���Ҫ���ܣ�"
	response.end
end if
if xz=1 then
	d="ֻ��"&mymp&"���ɵ��ӷ��ɽ���"
	k=mymp&"�صأ��Ǳ�����Ա�ô��ߣ�ɱ�޺գ�����"
else
	d="�κ��˶����Խ���"
	k=mymp&"���������㽻�������ѣ���ӭ���ǹ��١�"
end if
rs.open "select * from r where a='"&mymp&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ɵľ�������δ�����������޸ģ�');window.close();}</script>"
	Response.End
end if
if rs("c")=c and rs("f")=f and rs("g")=g and rs("h")=h and rs("i")=i and rs("j")=j then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�иı��κ����ã�����Ҫ�޸ģ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
conn.execute "update r set c="&c&",d='"&d&"',f="&f&",g="&g&",h="&h&",i="&i&",j="&j&",k='"&k&"' where a='"&mymp&"'"
rs.open "select * from r order by id",conn,2,2
application.lock()
roomming=""
do while Not rs.Eof
	rooming=rooming&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&";"
	rs.MoveNext
loop
application("aqjh_room")=rooming
application.unlock()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../hc3w_Admin/setup.css">
<title>���ɾ����������</title></head>
<body text="#FFFFFF" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
  <font color="#000000" size="2">
  <p>���������ɾ�����</p>
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
        ��ϲ��<font color=red><b><%=mymp%></b></font>�����������޸���ɡ�</font><br>
        <br>
        <input  onClick="javascript:location.href = 'javascript:history.go(-1)';" value="�� ��" type=button name="back">
        -------- 
        <input  onClick="javascript:window.close()" value="�� ��" type=button name="close">
      </div>
    </td>
  </tr>
</table>