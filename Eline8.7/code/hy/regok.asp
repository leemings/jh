<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
pass= md5(trim(Request.form("pass")))
hygrade=int(abs(Request.form("hygrade")))
hymonth=int(abs(Request.form("hymonth")))
money=abs(int(Request.form("money")))
'if not isnumeric(hygrade) then 
if hygrade<>1 and hygrade<>2 and hygrade<>3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ա�ȼ������');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if hymonth<1 or hymonth>12 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ա����ʱ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM q WHERE a='"&session("sjjh_name")&"'",conn,2,2
If not(Rs.Bof OR Rs.Eof) Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ������������Ա����δ����\nϵͳ�ܾ��������룡��');location.href = 'javascript:history.go(-1)';</script>"
			Response.End
end if
rs.close
rs.open "SELECT * FROM �û� WHERE ����='"&session("sjjh_name")&"' and ����='"&pass&"'",conn,1,2
If Rs.Bof OR Rs.Eof Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ���������벻��ȷ��');location.href = 'javascript:history.go(-1)';</script>"
			Response.End
end if
oldhy=rs("��Ա�ȼ�")
if rs("��Ա�ȼ�")<>hygrade and oldhy<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ��ǻ�Ա��ֻ�ܰ�����ʱ����\nҲ����˵����ֻ������["&oldhy&"]��Ա\nҪ����������վ����ϵ����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
'if rs("��Ա�ȼ�")>0 then
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ��ǻ�Ա������վ����ϵ\n�����Ա���ѣ���');location.href = 'javascript:history.go(-1)';</script>"
'	Response.End
'end if
if rs("��Ա�ȼ�")>1 then
	hydate=rs("��Ա����")
	
else
	hydate=date()
end if
n=Year(hydate)
y=Month(hydate)
r=Day(hydate)
y=y+hymonth
if y>12 then 
	y=y-12
	n=n+1
end if
jy=0


if hygrade=1 then jy=31250
if hygrade=2 then jy=90000
if hygrade=3 then jy=250000

if oldhy=1 then oldjy=31250
if oldhy=2 then oldjy=90000
if oldhy=3 then oldjy=250000
if oldhy=4 then oldjy=490000

sj=n & "-" & y & "-" & r
rs("��Ա�ȼ�")=hygrade
rs("��Ա����")=cdate(sj)
rs("allvalue")=rs("allvalue")+jy-oldjy
rs.Update
conn.execute "insert into q(a,b,d,e,f) values ('"& session("sjjh_name") &"',date(),'δ��','��Ա�ȼ���"&hygrade&",��Աʱ�䣺"&hymonth&"����,��"&money&"Ԫ!','"&userip&"')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ����ϲ�����Ļ�Ա�������,��Ա��ֹ����Ϊ��"&sj&"\n�����˳��������½��룡\n������"&date()+7&"��ǰ��ȡ�ﵽ��Ա����������һ�ܺ�ϵͳ����ɾ������ID��');location.href = '../chat/hy.asp';</script>"
Response.End 
%>