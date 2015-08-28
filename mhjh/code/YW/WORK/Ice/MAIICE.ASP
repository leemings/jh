<%
my=session("yx8_mhjh_username")
if my=""  then %>
<script language="vbscript">
  alert("您已经超时与服务器断开联接，请重新登陆！")
</script>
<%
response.end
end if
if Session("icets")<=0 then%>
<script language="vbscript">
  alert("你一桶冰也没有采，你拿什么卖钱？？？")
  </script>
<%
response.end
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+"& Session("icets")*20000 &" where 姓名='"& my &"'"
conn.execute sql
Response.Write "<script Language=Javascript>alert('您采冰："& session("icets") &"桶,卖了"& session("icets")*20000 &"两白银！');</script>"
Session("icets")=0
Session("icejl1")=0
Session("cbsj")=true
set rs=nothing	
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
