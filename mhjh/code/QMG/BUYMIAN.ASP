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
sql="select 积分 from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是剑侠中人，不能卖面"
conn.close
response.end
else
jishu=rs("积分")
sql="select * from 巧面 where 拥有者='" & name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "你没有面啊，是不是想骗我，快去面馆作面吧。"
location.href = "qmg.htm"
</script>
<%     connt.close
response.end
else
mw=rs("美味度")
tl=mw*8
yin=mw*3
sql="insert into 卖面(拥有者,体力,售价) values ('"& name &"',"& tl &","& yin &")"
connt.execute(sql)
connt.execute("delete * from 巧面 where 拥有者='" & name & "'" )
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+"& yin &"  where 姓名=' " & name & " ' "
conn.execute(sql)
conn.execute("update 用户 set 积分=积分+10 where 姓名='"&name&"'")
Response.write"你把你亲手做的面卖给了巧面馆，得到银子" & yin& "两"
set rs=nothing
conn.close
connt.close%>
<%end if%>
<%end if%>
