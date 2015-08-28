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
  alert("你一块铁也没有炼出来，你拿什么来保存呀？？？")
</script>
<%
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr		
	sql="select * from 物品 where 名称='铁块' and 所有者='" & my & "'"
			Set Rs=conn.Execute(sql)
			If Rs.Bof OR Rs.Eof then
				sql="insert into 物品(名称,所有者,属性,内力,体力,数量,价格) values ('铁块','"&my&"','药品',0,0,"& Session("tiets") &",50000)"
				conn.execute sql
				sql="update 物品 set 数量=数量-2 where 名称='矿石' and 所有者='"&my&"'"
				conn.execute sql
			else
				sql="update 物品 set 价格=50000,数量=数量+" & Session("tiets") & " where 名称='铁块' and 所有者='"&my&"'"
				conn.execute sql
				sql="update 物品 set 数量=数量-2 where 名称='矿石' and 所有者='"&my&"'"
				conn.execute sql
			end if
sql="select * from 物品 where 名称='铁块' and 所有者='" & my & "'"
Set Rs=conn.Execute(sql)
Session("tietsl")=rs("数量")
Response.Write "<script Language=Javascript>alert('您炼出铁块："& session("tiets") &"块,现拥有铁块"& rs("数量") &"块，多多努力！');</script>"
Session("tiets")=0
Session("tiesj")=true
set rs=nothing	
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
