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
rs.open "SELECT id FROM w WHERE b=" & aqjh_id & " and i>=10 and a='冰水'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的冰水不够10个不能操作！');</script>"
	response.end
end if
iceid=rs("id")
rs.close
rs.open "SELECT id FROM w WHERE b=" & aqjh_id & " and i>=4 and a='大草鱼'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的大草鱼不够4个不能操作！');</script>"
	response.end
end if
cyuid=rs("id")
conn.Execute "update w set i=i-15 WHERE id="&iceid
conn.Execute "update w set i=i-4 WHERE id="&cyuid
rs.close
rs.open "select * from x where a='三花聚顶' and b="&aqjh_id
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into x(a,b,c,d,e) values ('三花聚顶'," & aqjh_id &",1,'武功',5000)"
else
	id=rs("id")
	conn.execute "update x set e=5000,c=c+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("您三花聚顶修练完成！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>