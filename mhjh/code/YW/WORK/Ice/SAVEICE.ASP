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
  alert("你一桶冰也没有采，你拿什么来保存呀？？？")
</script>
<%
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
			sql="select * from 物品 where 名称='冰水' and 所有者='" & my & "'"
			Set Rs=conn.Execute(sql)
			If Rs.Bof OR Rs.Eof then
				sql="insert into 物品(名称,所有者,属性,内力,体力,数量,价格) values ('冰水','"&my&"','药品',10000,10000,"& Session("icets") &",20000)"
				conn.execute sql
			else
				sql="update 物品 set 价格=20000,数量=数量+" & Session("icets") & " where 名称='冰水' and 所有者='"&my&"'"
				conn.execute sql
			end if
sql="select * from 物品 where 名称='冰水' and 所有者='" & my & "'"
Set Rs=conn.Execute(sql)
Session("icesl")=rs("数量")
Response.Write "<script Language=Javascript>alert('您采冰："& session("icets") &"桶,现拥有冰水"& rs("数量") &"桶，多多努力！');</script>"
Session("icets")=0
Session("icejl1")=0
Session("cbsj")=true
set rs=nothing	
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
