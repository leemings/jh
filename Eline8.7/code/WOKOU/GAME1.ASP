<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
If Session("usepro") = true Then
win=request("win")
my=Session("sjjh_name")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if win<>0 then
rs.open "select * from 用户 where 姓名='"&my&"'",conn
conn.execute "update 用户 set 内力=内力+10,体力=体力+10,攻击=攻击+10,防御=防御+10,银两=银两+200  where 姓名='"&my&"'"
rs.close
set rs=nothing	
set conn=nothing
Session("usepro")= false 

says="<font color=#ff0000><b>【抗击倭寇】" & sjjh_name & "</b>大侠，你抗击倭寇有功，『E线江湖』奖你内力、体力、攻击、防御各10点，银两200两！继续努力哦：）</font>"			'聊天数据
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub

Response.Write "<script Language=Javascript>alert('大侠，你打击倭寇有功，盟主奖内力、体力、攻击、防御各10点，银两200两！');</script>"
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