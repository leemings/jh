<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Application("aqjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有鬼？！');</script>"
	response.end
end if
id=LCase(trim(Request.QueryString("id")))
tsk=int(abs(clng(Request.QueryString("tsk"))))
tempjs=int(abs(clng(Application("aqjh_kl"))))
if tempjs<>tsk then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where 状态='僵尸王' and 身份='僵尸王' order by 姓名 ",conn
sf=rs("身份")
attackname=rs("姓名")
tl=rs("体力")
if aqjh_grade>1 then
	Response.Write "<script language=JavaScript>{alert('提示：管理2级以上者无资格打僵尸王');window.close();}</script>"
	Response.End 
end if
Application.Lock
Application("aqjh_kl")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select 体力,w6 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("体力")<-100000 then
	conn.execute "update 用户 set 状态='僵尸王',身份='僵尸王',事件原因='僵尸王|"&fn1&"' where 姓名='" & aqjh_name &"'"
    conn.execute "update 用户 set 体力=5000,内力=5000,事件原因='无',状态='正常',身份='弟子' where 姓名='" & attackname &"'"
  	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'僵尸王','因勇打僵尸所牺牲','人命')"
	kl="<img src='pic/kl.gif'>["&aqjh_name&"]太衰了，为了跟僵尸王" & attackname & "搏斗结果自己反而被僵尸王"&id&"咬死,["&aqjh_name&"]变成了僵尸王." & attackname & "正式复活各位大虾恭喜他吧" & attackname & "被困在僵尸库好久了..，………"
	call boot(aqjh_name,"僵尸王，操作者：僵尸王，打我呀找死！")
else
dim js(10)
js(0) ="骨头"
js(1) ="蓝宝石"
js(2) ="红宝石"
js(3) ="僵尸血"
js(4) ="僵尸牙"
js(5) ="水晶石"
js(6) ="绿宝石"
js(7) ="硫磺"
js(8) ="硝酸"
js(9)="木炭"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w6"),js(myxy),2)
conn.execute "update 用户 set 银两=银两+"& tempjs*1 &",w6='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="我call,英雄，真是英雄，["&aqjh_name&"]与僵尸王" & attackname & "大打出手……僵尸王终于倒下了，而他自己也伤的不轻，官府奖励<img src='img/251.gif'>"&tempjs*1&"两！在打败僵尸王后，"&aqjh_name&"得到了[<b><font color=red>"&js(myxy)&"</font></b>]2个……"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【打僵尸王】</b></font>"&kl
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
%>