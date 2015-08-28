<%
my=session("yx8_mhjh_username")
if my=""  then %>
<script language="vbscript">
  alert("您已经超时与服务器断开联接，请重新登陆！")
</script>
<%
response.end
end if
if Session("minets")>0 then
Session("minets")=Session("minets")-1
end if
if Session("minets")<=0 then%>
<script language="vbscript">
  alert("你一块矿石也没有，你拿什么卖钱？？？")
  </script>
<%
response.end
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+"& Session("minets")*30000 &" where 姓名='"& my &"'"
conn.execute sql
Response.Write "<script Language=Javascript>alert('您采矿石："& session("minets") &"块,卖了"& session("minets")*30000 &"两白银！');</script>"
Session("minets")=0
Session("minesj")=true
set rs=nothing	
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>