<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=15 and 名称='冰水'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("您冰水不够15个，所以不能操作！")
window.close()
</script><%
response.end
end if
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=2 and 名称='小鲤鱼'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("您小鲤鱼不够2条，所以不能操作！")
window.close()
</script><%
response.end
end if
sql="update 物品 set 数量=数量-15 WHERE 所有者='" & name & "' and 名称='冰水'  "
Set Rs=conn.Execute(sql)
sql="update 物品 set 数量=数量-2 WHERE 所有者='" & name & "' and 名称='小鲤鱼'  "
Set Rs=conn.Execute(sql)
sql="select * from 物品 where 名称='九花玉露丸' and 所有者='" & name & "' and 属性='药品'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
sql="insert into 物品(名称,所有者,数量,属性,特效,体力,防御,攻击,内力,寄售,装备) values ('九花玉露丸','" & name &"',1,'药品','无',0,0,0,1000000,False,False)"
conn.execute sql
else
sql="update 物品 set 数量=数量+1 where 名称='九花玉露丸' and 所有者='"&name&"'"
conn.execute sql
end if
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">alert("您九花玉露丸修练完成！吃下能长内力100万！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>