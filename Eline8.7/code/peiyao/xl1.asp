<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT id FROM w WHERE b=" & sjjh_id & " and i>=5 and a='��ˮ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����ı�ˮ����5�����ܲ�����');</script>"
	response.end
end if
iceid=rs("id")
conn.Execute "update w set i=i-5 WHERE id="&iceid
rs.close
rs.open ="select id from x where a='������' and b="&sjjh_id
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into x(a,b,c,d,e) values ('������',"&sjjh_id&",1,'����',1800)"
else
	id=rs("id")
	conn.execute "update x set e=1500,c=c+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("��������������ɣ�")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>
