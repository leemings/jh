<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if name="" then Response.Redirect "../error.asp?id=016"
%>
<!--#include file="dadata.asp"-->
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select * from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是剑侠中人，不能吃面"
conn.close
response.end
else
js=rs("积分")
sql="select * from 巧面 where 拥有者='" & name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "你没有面啊，是不是想骗我，快去面馆作面吧。"
location.href = "qmg.htm"
</script>
<% 
connt.close
response.end
else
mw=rs("美味度")
zjtl=int(mw*5)
conn.execute("update 用户 set 体力=体力+" & zjtl & "  where 姓名='"&name&"'")
conn.execute("update 用户 set 积分=积分+1 where 姓名='"&name&"'")
connt.execute("delete * from 巧面 where 拥有者='" & name & "'" )
Response.write"你一口气把香气扑鼻的面吃下去。增加体力" & zjtl& ""
rs.close
conn.close
connt.close
end if
end if
%>
