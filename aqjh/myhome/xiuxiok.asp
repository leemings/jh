<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
act=clng(trim(request("act")))
if act<>1 and act<>2 then act=1
if act=1 then
	sj=clng(trim(Request("sj")))
	if sj<1 or sj>72 then
		Response.Write "<script Language=Javascript>alert('��ʾ��ʱ��ѡ�����!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
else
	if Hour(time())<21 or Hour(time())>23 then
		Response.Write "<script Language=Javascript>alert('��ʾ���ѻ�������Ϣʱ��Ϊÿ��21-23��,ÿ��8Сʱ!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "select h_�;ö� from house  where "& session("myroom") &"='"& aqjh_name &"'", conn, 1, 1
if rs.Eof and rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('��ʾ:û�й������ǲ�����Ϣ�ģ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("h_�;ö�")<10 then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
   	Response.Write "<script language=JavaScript>{alert('��ʾ:�������������,���޸������ɣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
if act=1 then
	rs.open "SELECT ״̬,��¼,�¼�ԭ��,���� from [�û�] WHERE ����='" & aqjh_name & "'",conn,1,3
	if rs("״̬")<>"����" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�������ڵ�״̬���������ܽ�����Ϣ!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	rs("״̬")="��"
	rs("��¼")=DateAdd("h", sj, now())
	rs("�¼�ԭ��")="�����Լ�������������Ϣ��:"& sj &"Сʱ,����ʱ������:"& DateAdd("h", sj, now()) 
	rs("����")=rs("����")+sj*80000
	rs.update
	rs.close
	show="��ʾ����Ϣ�ɹ�,��������:"& sj *80000 &"��\n,��ȷ�����˳�����,����"& sj &"Сʱ���¼����!"
else
	rs.open "SELECT ��ż from [�û�] WHERE ����='" & aqjh_name & "'",conn,1,1
	mywife=rs("��ż")
	rs.close
	rs.open "SELECT ״̬ from [�û�] WHERE ����='" & mywife & "'",conn,1,1
	if rs.Eof or rs.Bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ���Ҳ��������ż����!');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	wifezt=rs("״̬")
	rs.close
	rs.open "SELECT ��ż,״̬,��¼,�¼�ԭ��,���� from [�û�] WHERE ����='" & aqjh_name & "'",conn,1,3
	rs("״̬")="��"
	rs("��¼")=DateAdd("h", 8, now())
	rs("�¼�ԭ��")="���Ѿ�����,�ǻؼ���Ϣ��ʱ����,ÿ�챣֤8Сʱ��Ϣ�������,��������100���!����ʱ������:"& DateAdd("h", 8, now())
	rs("����")=rs("����")+15000
	rs.update
	rs.close
	'�����Ժ�Ҫ�ж�С���ӹ���,��ȷ������.
	show="��ʾ���ѻ������Ϣ�ɹ�,��������:100���\n,��ȷ�����˳�����,����8Сʱ���¼����!"
end if
rs.Open "select h_�;ö� from house  where "& session("myroom") &"='"& aqjh_name &"'", conn, 1, 3
rs("h_�;ö�")=rs("h_�;ö�")-1
rs.update
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('"& show &"');top.location.href = '../exit.asp';</script>"
response.end
%>