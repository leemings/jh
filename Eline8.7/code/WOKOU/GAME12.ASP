<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
If Session("usepro") = true Then
win=request("win")
my=Session("sjjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if win<>0 then
rs.open "select * from 用户 where 姓名='"&my&"'",conn
conn.execute "update 用户 set 内力=内力+1000,体力=体力+1000,攻击=攻击+1000,防御=防御+1000,银两=银两+2000 where 姓名='"&my&"'"
Session("usepro")= false 
rs.close
conn.close
set rs=nothing
set conn=nothing
Session("usepro")= false 
Response.Write "<script Language=Javascript>alert('大侠，你打击倭寇有功，『E线江湖』奖内力、体力、攻击、防御各1000点，银两2000两！');</script>"
response.end
%>
<script language="vbscript">
window.close()
</script> 
<%
else 
rs.open "select * from 用户 where 姓名='"&my&"'",conn
conn.execute "update 用户 set 内力=内力-300,体力=体力-300,攻击=攻击-300,防御=防御-300,银两=银两-20 where 姓名='"&my&"'"
Session("usepro")= false 
rs.close
conn.close


ff="<font color=#ff0000>消息</font>"&my&"大侠打击倭寇有功：『E线江湖』奖内力、体力、攻击、防御各1000点，银两2000两！"	'聊天数据






set rs=nothing
Response.Write "<script Language=Javascript>alert('菜鸟，平时不练功，现在打败了，还敢来见我，罚你体力和内力300，银两20个！');</script>"
response.end
%>
<script language="vbscript">
window.close()
</script>	
<%
end if
rs.close
conn.close
set rs=nothing
Session("usepro")= false 
else 
response.write "请通过正常途径来打倭寇." 
response.end 
end if
%>
