<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="../../showchat.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
lx=request("lx")
name=request("name")
sl=request("sl")
danjia=request("danjia")
select case lx
case "��Ʒ"
dang="w6"
case "��ҩ"
dang="w8"
case "ħ��"
dang="w9"
case "����"
dang="w10"
case "װ��"
dang="w3"
case else
Response.Write "<script language=javascript>alert('�Բ������ͳ���,��������?'); window.close();</script>"
Response.End
end select
if name="" then
Response.Write "<script language=javascript>alert('�Բ�����Ʒ������!'); window.close();</script>"
Response.End
end if
if not isnumeric(sl) or not isnumeric(danjia) then
Response.Write "<script language=javascript>alert('�Բ��������򵥼۳���!'); window.close();</script>"
Response.End
end if
if sl<1 then
Response.Write "<script language=javascript>alert('�Բ�����ʲô�������������ʲô������ȥ��'); window.close();</script>"
Response.End
end if
if danjia<10000 then
Response.Write "<script language=javascript>alert('�Բ��𣬻�����10000���������ǲ������أ�����������ˣ��������ͳ�ȥ�������˾�ȥ���������ͣ�'); window.close();</script>"
Response.End
end if
if danjia>1000000000 then
Response.Write "<script language=javascript>alert('�Բ��𣬶�����10���ˣ�˭�������'); window.close();</script>"
Response.End
end if
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from �û� where ����='"&aqjh_name&"'"
rst.open sqlstr,conn
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('�Բ����㲻�ǽ������ˣ�');window.close();</script>"
Response.End
end if
dangwp=rst(dang)
if IsNull(dangwp) then
Response.Write "<script language=javascript>alert('�Բ�����û���������͵���Ʒ����');window.close();</script>"
Response.End
end if
if instr(dangwp,name)=0 then
Response.Write "<script language=javascript>alert('�Բ�����û��������Ʒ����');window.close();</script>"
Response.End
end if
data1=split(dangwp,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
	if data3(0)=name then
		if sl>data3(1) then
			Response.Write "<script language=javascript>alert('�Բ��������Ʒ������������');window.close();</script>"
			Response.End
		end if
	'ɾ����Ʒ
	wpsl=sl
	mydang=abate(dangwp,name,wpsl)
	conn.execute "update �û� set "&dang&"='"&mydang&"' where  ����='"&aqjh_name&"'"
	Set connb = Server.CreateObject("ADODB.Connection")
	connb.Open Application("aqjh_usermdb")
	Set rst= Server.CreateObject("ADODB.Recordset")
	sql="select * from ���� where ����='"&name&"' and ������='"&aqjh_name&"'"
	Set rst=connb.Execute(sql) 
	if rst.EOF or rst.BOF then
		sqlstr1="insert into ����(����,����,�۸�,����,������) values ('"&name&"','"&dang&"',"&danjia&","&wpsl&",'"&aqjh_name&"')"
	else
		sqlstr1="update ���� set ����=����+"&wpsl&",�۸�="&danjia&" where ����='"&name&"' and ������='"&aqjh_name&"'"
	end if
	connb.Execute(sqlstr1)
	says="<font color=red>���г���Ϣ��</font><font color=blue>"&aqjh_name&"���Լ������"&wpsl&"��"&name&"�ڽ����������г������ˣ�����Ϊ"&danjia&"���ӣ���Ҫ�Ĵ����Ͻ�ȥ����Ŷ��������û����ˣ�</font>"
	call showchat(says)
	exit for
	end if
next
%>
<head><title>���ɼ���</title><link rel="stylesheet" href="../../css.css"><script language=javascript>setTimeout("location.replace('sale.asp');",3000);</script></head>
<body background="../../bg.gif" oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td bgcolor=#005b00><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
</tr></table>
<p align=center><font color="#CC0000" size="4" face="��Բ"><b><%=Application("aqjh_chatroomname")%>���ɼ���</b></font></p>
<div align="center">��Ļ������������ˣ�����������ǽ����Ͻ�Ǯ���㣬ˡ������֪ͨ�ˣ��������10% </div>
<p align=center><a href="javascript:location.replace('sale.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>