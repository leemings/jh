<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
ndate=Day(date())
if ndate<21 or ndate>27 then
		Response.Write "<script Language=Javascript>alert('��ʾ��ͶƱʱ��Ϊÿ��21��27�ţ�����ʱ�䲻���Խ��У�');window.close();</script>"
		response.end
end if
id=clng(Request("id"))
if Session("aqjh_jhdj")<20 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǽ������֣�����20�����ܲ���ͶƱ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ա�,��ż,����,������,ʦ�� from [�û�] where ����='" & Session("aqjh_name") & "'",conn,1,1
mysex=rs("�Ա�")
po=rs("��ż")
qr=rs("����")
jsr=rs("������")
sf=rs("ʦ��")
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,1
if instr(rs("p_ͶƱ��"),","& aqjh_name &"|")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������涨��ÿ��ֻ����ͶһƱ�����Ѿ�Ͷ���ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.Open "select * from sai where s_id="& id, conn, 1,3
if mysex=rs("s_�Ա�") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������涨ֻ��������ͶƱ�����ж�Ů����ͶƱ��Ů���н���ͶƱ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
toname=rs("s_����")
if po=toname or qr=toname or jsr=toname or sf=toname then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����ż�����ˡ������ˡ�ʦ���˲����Խ���ͶƱ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs("s_Ʊ��")=rs("s_Ʊ��")+1
rs.update
rs.close
rs.Open "select * from pool", conn, 1,3
rs("p_ͶƱ��")=rs("p_ͶƱ��") & toname & "," & aqjh_name & "|"
rs("p_ͶƱ����")=rs("p_ͶƱ����")+1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from [�û�] where ����='" & Session("aqjh_name") & "'",conn,1,3
rs("����")=rs("����")+10000
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>����Ϣ��##�ڽ����������Ͷ�������һƱ��ϵͳ����1������������������ڽ����У�Ҫ������ѡ�Ŀ�ȥ������<font class=t>(" & time() & ")</font></font>"
act="��Ϣ"
towho="���"
addwordcolor="660099"
saycolor="008888"
call addsay(Session("inroom"),aqjh_name,"��",towho,addwordcolor,saycolor,says,act,0)
Response.Write "<script Language=Javascript>alert('��л��Ͷ������һƱ��ϵͳ�Զ�Ϊ������һ������Ϊ������');location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>