<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
name=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id FROM w WHERE b=" & aqjh_id & " and i>=8 and a='��ʯ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ŀ�ʯ����8�����ܲ�����');</script>"
	response.end
end if
kuangid=rs("id")
rs.close
rs.open "SELECT id FROM w WHERE b=" & aqjh_id & " and i>=8 and a='��ˮ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����ı�ˮ����8�����ܲ�����');</script>"
	response.end
end if
iceid=rs("id")
rs.close
rs.open "SELECT id FROM w WHERE b=" & aqjh_id & " and i>=1 and a='�����'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ĵ���㲻��1�����ܲ�����');</script>"
	response.end
end if
caoyuid=rs("id")
conn.Execute "update w set i=i-8 WHERE id="&kuangid
conn.Execute "update w set i=i-8 WHERE id="&iceid
conn.Execute "update w set i=i-1 WHERE id="&caoyuid
rs.close
rs.open "select * from x where a='�ָ�ҩ��' and b="&aqjh_id,conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into x(a,b,c,d,e) values ('�ָ�ҩ��'," & aqjh_id &",1,'���ͨ��',4000)"
else
	id=rs("id")
	conn.execute "update x set e=4000,c=c+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("���Ļָ�ҩ��������ɣ�")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>