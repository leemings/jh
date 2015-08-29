<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name=""  then Response.Redirect "../error.asp?id=440"
id=request("id")
%><!--#include file="dadata.asp"-->
<%
sql="select * from 用户 where 姓名='"&sjjh_name&"'" 
set rs=conn.execute(sql)
if rs.eof or rs.bof then
conn.close
Response.Write "<script language=javascript>alert('你没有门派，不能订购此酒宴');history.back();</script>"
response.end
else
mypai=rs("门派")
sql="SELECT * FROM 宴会 where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("宴会名")
yin=rs("售价")
lx=rs("类型")
nl=rs("内力")
tl=rs("体力")
jb=rs("金币")
sl=rs("数量")%>

<%
sql="select * from 用户 where 姓名='"&sjjh_name&"'"
rs=conn.execute(sql)
if yin<=rs("银两") then
sql="update 用户 set 银两=银两-" & yin & " where 姓名='"&sjjh_name&"'"
rs=conn.execute(sql)%>
<%
sql="select * from 宴会列表 where 宴会名='" & wu & "' and 拥有者='" & sjjh_name & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into 宴会列表(宴会名,拥有者,类型,内力,体力,金币,数量,售价,门派,时间) values ('"&wu&"','"&sjjh_name&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&mypai&"',now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
says="<font color=red>【好消息】</font><b><font color=green>"&sjjh_name&"在江湖大酒店的"&wu&"厅为"&mypai&"举行<font color=red>※"&lx&"※</font>宴会，"&mypai&"的都快去呀，晚了就没的吃了。。。</font></b>"
says=replace(says,"'","\'")
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
Response.Write "<script language=javascript>alert('恭喜你，你的酒宴定桌成功，去告诉你的朋友吧！');history.back();</script>"			
Response.Redirect "jd.asp"
else

connt.close				
Response.Write "<script language=javascript>alert('不能定酒宴，原因：你已定了同样规格的酒席！');history.back();</script>"
response.end
				
end if
else
connt.close
Response.Write "<script language=javascript>alert('不能定酒宴，原因：你的银两不够了');history.back();</script>"
response.end
end if
rs.close
set rs=nothing
end if
%>