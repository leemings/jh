<%
my=session("yx8_mhjh_username")
if my=""  then %>
<script language="vbscript">
  alert("您已经超时与服务器断开联接，请重新登陆！")
</script>
<%
response.end
end if
if Session("tiets")>0 then
Session("tiets")=Session("tiets")-1
end if
if Session("tiets")<=0 then%>
<script language="vbscript">
  alert("你一块铁块也没有，你拿什么卖钱？？？")
  </script>
<%
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="update 用户 set 银两=银两+"& Session("tiets")*50000 &" where 姓名='"& my &"'"
conn.execute sql

if Session("tiesl")<=0 then
Session("tiesl")=0
end if

sql="update 物品 set 数量="& Session("tiesl") &" where 名称='矿石' and 所有者='"&my&"'"
conn.execute sql
%>
Response.Write "<script Language=Javascript>alert('您炼铁块：<%=session("tiets")%>块,卖了<%=session("tiets")*50000%>两白银！');</script>"
<%
Session("tiets")=0
Session("tiesj")=true
set rs=nothing	
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>


