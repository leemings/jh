<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="../../showchat.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
cmid=Request.QueryString("id")
sl=Request.Form("sl")
if not (isnumeric(cmid) and isnumeric(sl)) then
Response.Write "<script language=javascript>alert('�Բ���������ʹ�����֣�'); history.back();</script>"
Response.End
end if
sl=clng(sl)
Set connb = Server.CreateObject("ADODB.Connection")
connb.Open application("aqjh_usermdb")
Set rst= Server.CreateObject("ADODB.Recordset")
sqlstr="select * from ���� where id="&cmid
rst.open sqlstr,connb
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('�Բ����г���û��������Ʒ���ۻ����Ѿ����꣡'); history.back();</script>"
Response.End
end if
syz=rst("������")
danjia=rst("�۸�")
name=rst("����")
lx=rst("����")
shul=rst("����")
yin=sl*danjia
if sl<1 then
Response.Write "<script language=javascript>alert('�Բ�����ʲô�������������ʲô����ƭ�Ҷ�����������ȥ��'); history.back();</script>"
Response.End
end if
'shulΪ����,slΪ�������
if sl>shul then
Response.Write "<script language=javascript>alert('�Բ���û��"&sl&"��ô��"&name&"��Ŷ��Ŀǰ��ʣ��"&shul&"����'); history.back();</script>"
Response.End
end if
rst.Close
set conn=server.CreateObject("adodb.connection")
conn.Open Application("aqjh_usermdb")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ���� from �û� where ����='"&aqjh_name&"' "
rst.open sqlstr,conn
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('�Բ������Ѿ��Ͽ����ӣ�'); history.back();</script>"
Response.End
end if
if rst("����")<yin then
Response.Write "<script language=javascript>alert('�Բ����ܹ���Ҫ����"&yin&"����Ǯ����Ŷ��'); history.back();</script>"
Response.End
end if
rst.Close
sqlstr="select * from �û� where ����='"&aqjh_name&"'"
rst.open sqlstr,conn
dangwp=rst(lx)
'������Ʒ
addsl=sl
zstemp=add(dangwp,name,addsl)
conn.execute "update �û� set "&lx&"='"&zstemp&"' where  ����='"&aqjh_name&"'"
if syz<>aqjh_name then
conn.execute "update �û� set ����=����-"&yin&" where  ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����+"&yin&" where  ����='"&syz&"'"
says="<font color=red>���г���Ϣ��</font><font color=blue>"&aqjh_name&"������"&yin&"�����������г��ﹺ����"&addsl&"��"&name&"������ͦ���˵�Ŷ��</font>"
else
says="<font color=red>���г���Ϣ��</font><font color=blue>"&aqjh_name&"�������г�ߺ���˰���,����û����["&name&"]���������,ֻ����̯�����ˣ�</font>"
end if
sqlstr1="update ���� set ����=����-"&addsl&" where ����='"&name&"' and ������='"&syz&"'"
connb.execute (sqlstr1)
sqlstr2="delete * from ���� where ����=0"
connb.execute (sqlstr2)
call showchat(says)
%>
<head><title>���ɼ���</title><link rel="stylesheet" href="../../css.css"><script language=javascript>setTimeout("location.replace('Market.asp');",2000);</script></head>
<body background="../../bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td bgcolor=#ffff00><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("aqjh_chatroomname")%>���ɼ���</b></font></p>
<p align=center>������Ʒ������Ҫ�����ֽ𣬹���ɹ�������ֱ�ӻ���</p>
<div align="center">���ǵĻ������պ��˰����������Ҫ�ˣ����������Ұ�������ȥ�ɣ� </div>
<p align=center><a href="javascript:location.replace('Market.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>