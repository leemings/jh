<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=10 and 名称='矿石'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("您矿石不够10个，所以不能操作！")
window.close()
</script><%
response.end
end if
sql="SELECT * FROM 物品 WHERE 所有者='" & name & "' and 数量>=5 and 名称='小鲤鱼'  "
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then %>
<script language="vbscript">alert("您小鲤鱼不够5条，所以不能操作！")
window.close()
</script><%
response.end
end if
sql="update 物品 set 数量=数量-10 WHERE 所有者='" & name & "' and 名称='矿石'  "
Set Rs=conn.Execute(sql)
sql="update 物品 set 数量=数量-2 WHERE 所有者='" & name & "' and 名称='小鲤鱼'  "
Set Rs=conn.Execute(sql)
sql="select * from 物品 where 名称='夺命金丹' and 所有者='" & name & "' and 属性='药品'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
sql="insert into 物品(名称,所有者,数量,属性,特效,体力,防御,攻击,内力,寄售,装备) values ('夺命金丹','" & name &"',1,'暗器','无',0,0,0,800000,False,False)"
conn.execute sql
else
sql="update 物品 set 数量=数量+1 where 名称='夺命金丹' and 所有者='"&name&"'"
conn.execute sql
end if
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">alert("您的夺命金丹修练完成！杀伤力达80万的超级暗器属于你了！")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>