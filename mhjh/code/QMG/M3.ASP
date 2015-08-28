<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../error.asp?id=016"
%>
<!--#include file="dadata.asp"-->
<%
sql="select * from 物品 where 类型='主料' and 拥有者='" & name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "你没有干面，快去菜场买吧。"
location.href = "ciachang.asp"
</script>
<%connt.close
response.end
else
mw=rs("美味度")
sql="select * from 巧面 where  拥有者='" & name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into 巧面 (拥有者,美味度) values ('"& name &"'," & mw & ")"
rs=connt.execute(sql)
sql="select * from 物品 where 类型='主料' and 拥有者='" & name & "'"
sql="delete * from 物品 where 类型='主料' and 拥有者='" & name & "'"
set rs=connt.execute(sql)
%>
<script language=vbscript>
MsgBox "你把面放进锅里，一碗好面就快在眼前了。"
location.href = "m4.htm"
</script>
<%else
sql="update 巧面 set 美味度=美味度+" &mw& " where 拥有者='" & name & "'"
rs=connt.execute(sql)
sql="select * from 物品 where 类型='主料' and 拥有者='" & name & "'"
sql="delete * from 物品 where 类型='主料' and 拥有者='" & name & "'"
set rs=connt.execute(sql)
%>
<script language=vbscript>
MsgBox "你把 面放进锅里，一碗好面就快在眼前了。"
location.href = "m4.htm"
</script><%
set rs=nothing
connt.close
end if
end if
%>
