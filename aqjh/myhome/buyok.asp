<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"

%>
<html>
<body background="IMAGES/bg.jpg">
<div align="center"><a href="javascript:history.go(-1);"><font color="#0000FF">�����ﷵ��</font></a></div>
<%
houseid=clng(trim(Request("houseid")))
if houseid<0 or houseid>5 then
	Response.Write "<script Language=Javascript>alert('��ʾ��ѡ�������д�������ѡ��');location.href = 'house.asp';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,1
if not(rs.Eof and rs.Bof) then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ������˷����ˣ�Ҫ����������ȥ����������');location.href = 'house.asp';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM housetype WHERE ht_���=" & houseid,conn,1,1
xlwp=rs("ht_1������")
if isnull(xlwp) or xlwp="" or instr(xlwp,"|")=0 or instr(xlwp,";")=0 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����������д��ܲ���\n���ҳ��򿪷�����ϵ!');</script>"
	response.end
end if
xadata=split(xlwp,";")
xadata1=UBound(xadata)
rs.close
rs.open "SELECT w6,�ȼ�,��Ա�ȼ�,����,��� from [�û�] WHERE ����='" & aqjh_name & "'",conn,1,3
duyao=rs("w6")
if isnull(duyao) or duyao="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����ʲô��ƷҲû��!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
mysql=""
yin=0
for ii=0 to xadata1
	xadata2=split(xadata(ii),"|")
	mysl=clng(xadata2(1))
	myxlwp=trim(xadata2(0))
	select case myxlwp
	case "�ȼ�"
		if rs("�ȼ�")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�����ĵȼ�δ��["& mysl &"]�������ܹ���!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
	case "���"
		if rs("���")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�����Ľ��û��["& mysl &"]�������ܹ���!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
		rs("���")=rs("���")-mysl
	case "����"
		yin=yin+mysl
		if rs("����")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ����������û��["& mysl &"]�������ܹ���!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
		rs("����")=rs("����")-mysl
	case "��Ա�ȼ�"
		if rs("��Ա�ȼ�")<mysl then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�����Ļ�Ա�ȼ�δ��["& mysl &"]�������ܹ���!');location.href = 'javascript:history.go(-1)';</script>"
			response.end
		end if
	case else
		if mywpsl(duyao,myxlwp)<mysl then
    		Set rs = Nothing
   			conn.Close
  			Set conn = Nothing
   			Response.Write "<script Language=Javascript>alert('��ʾ����Ʒ"& myxlwp &"�������㣿');</script>"
			response.end
		end if
		duyao=abate(duyao,myxlwp,mysl)
	end select
next
rs("w6")=duyao
rs.update
rs.close
rs.open "SELECT * FROM housetype WHERE ht_���=" & houseid,conn,1,3
h_name=rs("ht_С����")
h_nai=rs("ht_1���;ö�")
rs("ht_����")=rs("ht_����")+1
rs("ht_С���ʲ�")=rs("ht_С���ʲ�")+int(yin/10000)
rs.update
rs.close
rs.open "SELECT * FROM house",conn,1,3
rs.addnew
rs("h_С����")=h_name
rs("h_ӵ����")=aqjh_name
rs("h_ӵ����2")="��"
rs("h_����")=houseid
rs("h_����")=1
rs("h_�ι�")=0
rs("h_����ʱ��")=now()
rs("h_����ʱ��")=now()
rs("h_�ؼҴ���")=0
rs("h_�;ö�")=h_nai
rs("h_����Ǯ��")=0
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=""Javascript"">alert(""��ϲ�㣬���������Եļ��ˣ��Ժ����ȡ�����ӹ�������ֵĽ������"");window.close();</script>"
%></html>
