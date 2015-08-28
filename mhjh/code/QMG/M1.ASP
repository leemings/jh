<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
%>
<!--#include file="dadata.asp"-->
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select * from 用户 where 姓名='" & name & "' and 职业='厨师'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是剑侠厨师，不能做面"
conn.close
response.end
else
sql="select * from 物品 where 类型='汤料' and 拥有者='" & name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "你没有汤料，怎么做面啊，快到菜市场去买吧。"
location.href = "caichang.asp"
</script>
<%conn.close
response.end
else
mw=rs("美味度")
sql="select * from 巧面 where  拥有者='" & name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into 巧面 (拥有者,美味度) values ('"& name &"'," & mw & ")"
rs=connt.execute(sql)
sql="select * from 物品 where 类型='汤料' and 拥有者='" & name & "'"
sql="delete * from 物品 where 类型='汤料' and 拥有者='" & name & "'"
set rs=connt.execute(sql)
%>
<script language=vbscript>
MsgBox "你把鸡汤放进锅里，一碗好面就快在眼前了。"
location.href = "m2.htm"
</script>
<%else
sql="update 巧面 set 美味度=美味度+" &mw& " where 拥有者='" & name & "'"
rs=connt.execute(sql)
sql="select * from 物品 where 类型='汤料' and 拥有者='" & name & "'"
sql="delete * from 物品 where 类型='汤料' and 拥有者='" & name & "'"
set rs=connt.execute(sql)
%>
<script language=vbscript>
MsgBox "你把鸡汤放进锅里，一碗好面就快在眼前了。"
location.href = "m2.htm"
</script><%
set rs=nothing
connt.close
end if
end if
end if
%>
