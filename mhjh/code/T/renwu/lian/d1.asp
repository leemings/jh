<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=3 and 名称='桂花酒'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("您桂花酒不足3瓶，所以不能操作！")
window.close()
</script><%
response.end
end if
sql="update 物品 set 数量=数量-3 WHERE 所有者='" & name & "' and 名称='桂花酒'  "
Set Rs=conn.Execute(sql)
sql="select * from 物品 where 名称='寒冰刀' and 所有者='" & name & "' and 属性='武器'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
sql="insert into 物品(名称,所有者,数量,属性,特效,体力,防御,攻击,内力,寄售,装备) values ('寒冰刀','" & name &"',1,'武器','无',0,10000,2000,0,False,False)"
conn.execute sql
else
sql="update 物品 set 数量=数量+1 where 名称='寒冰刀' and 所有者='"&name&"'"
conn.execute sql
end if
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">alert("你获得了一把寒冰刀，至寒至刚的好武器！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>
