<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
ndate=Day(date())
if ndate<27 then
		Response.Write "<script Language=Javascript>alert('��ʾ�����ڻ�û�е��콱��ʱ�䣡');window.close();</script>"
		response.end
end if
myok=trim(lcase(cstr(server.HTMLEncode(Request("ok")))))
%>
<html><title>�������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<%act=clng(Request("act"))
if act<>1 and act<>2 and act<>3 and act<>4 then
	Response.Write "<script Language=Javascript>alert('��ʾ����������ȷ��');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,3
select case act
case 1
	mess="�������һ��"
	if trim(rs("p_�е�һ��"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��������,�㲢û�еý���');window.close();</script>"
		response.end
	end if
	if rs("p_�е�һ�콱")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��������,�����ٲ�����');window.close();</script>"
		response.end
	end if
	rs("p_�е�һ�콱")=true
case 2
	mess="������ڶ���"
	if trim(rs("p_�еڶ���"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��������,�㲢û�еý���');window.close();</script>"
		response.end
	end if
	if rs("p_�еڶ��콱")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��������,�����ٲ�����');window.close();</script>"
		response.end
	end if
	rs("p_�еڶ��콱")=true
case 3
	mess="Ů�����һ��"
	if trim(rs("p_Ů��һ��"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��������,�㲢û�еý���');window.close();</script>"
		response.end
	end if
	if rs("p_Ů��һ�콱")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��������,�����ٲ�����');window.close();</script>"
		response.end
	end if
	rs("p_Ů��һ�콱")=true
case 4
	mess="Ů����ڶ���"
	if trim(rs("p_Ů�ڶ���"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��������,�㲢û�еý���');window.close();</script>"
		response.end
	end if
	if rs("p_Ů�ڶ��콱")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��������,�����ٲ�����');window.close();</script>"
		response.end
	end if
	rs("p_Ů�ڶ��콱")=true
case else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ����������ȷ���ܲ�����');window.close();</script>"
		response.end
end select
'ȡ�����н����
money=rs("p_���")
'����ǵ�һ��Ϊ30%
if act=1 or act=3 then
	gold=int(money*0.2)
else
	gold=int(money*0.1)
end if
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
'���û���,���ӽ��
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��� from [�û�] where ����='" & Session("aqjh_name") & "'",conn,1,3
rs("���")=rs("���")+gold
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>����Ϣ��</font>["&session("aqjh_name")&"]<font color=red>�ڽ��������л��"& mess &",�õ���Ʒ���<font class=t6>"& gold &"</font>��,�µĽ���������������1�ž���,ϣ���������׼��!<font class=t>(" & time() & ")</font></font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('��ʾ�����ڽ����ȡ�ɹ�,�õ�["& gold &"]����ң�');window.close();</script>"
response.end
%>
</body></html>