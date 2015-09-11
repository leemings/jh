<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request.form("id"))
qingren=trim(request.form("name"))
my=aqjh_name
%><!--#include file="dadata.asp"-->
<%
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是江湖中人，不能订购酒宴"
conn.close
response.end
else
sql="select * from 用户 where 姓名='" & qingren & "' and 姓名<>'" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
conn.close
Response.Write "<script language=javascript>alert('江湖中没有你输入的名字!');history.back();</script>"
response.end
else
qingren=rs("姓名")
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
sql="select * from 用户 where 姓名='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("银两") then
sql="update 用户 set 银两=银两-" & yin & " where 姓名='" & my & "'"
rs=conn.execute(sql)%>
<%
sql="select * from 宴会列表 where 宴会名='" & wu & "' and 拥有者='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into 宴会列表(宴会名,拥有者,类型,内力,体力,金币,数量,售价,参加者,时间) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&qingren&"',now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
if Instr(Application("aqjh_useronlinename"&session("nowinroom"))," "&qingren&" ")<>0 then
says="<font color=red>【好消息】</font><font color=green>"&my&"在江湖大酒店的"&wu&"厅为"&qingren&"举行<font color=red>※"&lx&"※</font>宴会，"&qingren&"快去呀，呵呵..</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
call showchat(says)
end if			
Response.Write "<script language=javascript>alert('恭喜，你为"&qingren&"定的酒宴已经准备好了！');window.close();</script>"
response.end
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
end if
%>