<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then 
Response.Write "<script language=javascript>{alert('��ʾ����û�е�½���Ѿ���ʱ�Ͽ����ӣ������µ�½��');parent.history.go(-1);}</script>" 
Response.End 
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gwm="ǧ��Ы��"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb") 
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
tl=rs("����")
nl=rs("����")
fl=rs("����")
hydj=rs("��Ա�ȼ�")
dj=rs("�ȼ�")
select case hydj
	case 0
		jgsj=320
	case 1
		jgsj=260
	case 2
		jgsj=220
	case 3
		jgsj=180
	case 4
		jgsj=130
	case 5
		jgsj=100
        case 6
		jgsj=60
sj=DateDiff("s",rs("����ʱ��"),now())
	s=jgsj-sj
end select
if dj<50 and dj>100 then 
	            Response.Write "<script Language=Javascript>alert('��ʾ�����ĵȼ�̫�͡���������������������ǣ��ȼ�����50��С��100������');location.href = 'javascript:history.go(-1)';</script>"
	            Response.End
end if
if tl<0 then 
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ� ���ˡ���55555555�´�ע��㡣��');location.href = '../../../exit.asp';</script>"
	Response.End
end if
%>
<HTML>
<HEAD>
<TITLE>�λ�ħ��--<%=gwm%></TITLE>
<link href="../../dg/setup.css" rel=stylesheet type="text/css">
<body oncontextmenu=self.event.returnValue=false bgcolor=#000000>
<br><br><br>
<table align=center width="300">
<div align=center><%
if Minute(time())<40 or Minute(time())>=60 then
	Response.Write "<font color=00ffff face=����>�λ�ħ����δ����</font>"
else
	Response.Write "<font color=00ff00 face=����>�λ�ħ�翪ʼ����</font>"
end if
%><br><br>
<div align=center><img src=../pic/xz.gif width=160 height=128></div>
<p align=center> <a href=xz1.asp><font color=ffff00><b>��������</b></font></a> <a href=zhx.asp><font color=red><b>���ѹ���</b></font></a></p>

<p align=left><font color=cccccc size=2>˵����<font color=00ff00>ÿ��Сʱ���ʮ����</font>���������λ�ħ����ԣ�������������Ҫ����<font color=ffff00>50000</font>�㣬ע����Ҿ��ǵģ��Ǻǣ�����������<font color=ffff00>���ݹ�������Լ�����������ã�������������Ʒ�����֣���㣩����ҡ���Ƭ����ҩ���߼�װ��<b><b><font color=00ff00>1000</font></b></b>�㷨������Ҫע�������ѵĹ����������Ҳ���Դ�Ҫץ��ʱ��Ŷ��</font></p>
</table><br><br>
<table align=center width="500">
<div align=center><font color=cccccc><font color=00ffff>��ɱ��</font>��<font color=cccccc><%=rs("����")%><font> <font color=00ffff>��Ա�ȼ�</font>��<font color=cccccc><%=rs("��Ա�ȼ�")%><font>�� <font color=00ffff>ʱ������</font>��<font color=00ff00><b><%=jgsj%></b><font><font color=ffff00> ��</font> 
<% if s>0 then 
Response.Write "<font color=00ffff>����</font>��<font color=00ff00><b>"&s&"</b><font><font color=ffff00> ��</font></font></div>"
else
Response.Write "<font color=00ff00>������ɱ<font></font></div>"
end if %>
</table><br><br>
<table border="1" width=650 cellspacing="1" cellpadding="0" bordercolor="eeeeee" height="26" align=center>
<tr>
<td height=25 align=center><font color=cccccc>������</font></td>
<td height=25 align=center><font color=cccccc>����</font></td>
<td height=25 align=center><font color=cccccc>ɱ��</font></td>
<td height=25 align=center><font color=cccccc>����</font></td>
<td height=25 align=center><font color=cccccc>����</font></td>
<td height=25 align=center><font color=cccccc>�ȼ�</font></td>
<td height=25 align=center><font color=cccccc>״̬</font></td>
<td height=25 align=center><font color=cccccc>��ɱ��</font></td>
</tr>
<% 
rs.close
gwm="ǧ��Ы��"
rs.open "select * from ���� where ����='"&gwm&"'",conn,2,2
%>
<tr>
<td height=25 align=center><font color=00ff00><%=rs("����")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("����")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("ɱ��")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("����")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("����")%></font></td>
<td height=25 align=center><font color=cccccc><%=rs("�ȼ�")%></font></td>
<td height=25 align=center><font color=ffff00><%=rs("״̬")%></font></td>
<td height=25 align=center><font color=00ffff><%=rs("��ɱ��")%></font></td>
</tr>
<% rs.close %>
</table><br><br><br><br>
<table align=center width="600">
<div align=center><a href=index.asp><font color=ffff00>��������ҳ��</font></a> <a href="javascript:location.reload()"><font color=00ff00>��ˢ�±�ҳ��</font></a><br><br><br><font size="2">��Ȩ���� <font color="#FFFFFF">�����齭������</font></font> 
  <font color="#FF0000" size="2">��������</font><font size="2"> �޸�����  �뱣����Ȩ</font></div>









