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
rs.open "SELECT id FROM w WHERE b=" & sjjh_id & " and i>=8 and a='冰水'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的冰水不够8个不能操作！');</script>"
	response.end
end if
iceid=rs("id")
rs.close
rs.open "SELECT * FROM w WHERE b=" & sjjh_id & " and i>=8 and a='矿石'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的矿石不够8个不能操作！');</script>"
	response.end
end if
kuangid=rs("id")
rs.close
rs.open "SELECT id FROM w WHERE b=" & sjjh_id & " and i>=3 and a='大鲨鱼'  "
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的大鲨鱼不够3个不能操作！');</script>"
	response.end
end if
sayuid=rs("id")
rs.close
rs.open "SELECT id FROM w WHERE b=" & sjjh_id & " and i>=2 and a='老虎肉'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的老虎肉不够2个不能操作！');</script>"
	response.end
end if
laohuid=rs("id")
conn.Execute "update w set i=i-8 WHERE id="&iceid
conn.Execute "update w set i=i-8 WHERE id="&kuangid
conn.Execute "update w set i=i-3 WHERE id="&sayuid
conn.Execute "update w set i=i-2 WHERE id="&laohuid
rs.close
rs.open "select id from x where a='爆胎易筋丸' and b="&sjjh_id
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into x(a,b,c,d,e) values ('爆胎易筋丸'," & sjjh_id &",1,'武功',15000)"
else
	id=rs("id")
	conn.execute "update x set e=15000,c=c+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("您爆胎易筋丸修练完成！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>