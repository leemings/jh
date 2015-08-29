<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
types=Request.QueryString("typename")
id=Request.QueryString("id")
Select Case id
Case "1"
  classes="状元"
Case "2" 
   classes="榜眼"
Case "3"
   classes="探花"
End Select
if classes=null or types=null then
	Response.Write "<script language=JavaScript>{alert('提示：数据错误，想黑我的话滚');}</script>"
	Response.End 
end if
Select Case types
Case "gold"
  typeses="金榜"
Case "silver" 
   typeses="银榜"
Case "copper"
   typeses="铜榜"
End Select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 武功,攻击,防御,内力,性别,门派,身份,体力 from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是江湖中人。');window.close();}</script>"
	Response.End 
end if
mywg=rs("武功")
mygj=rs("攻击")
myfy=rs("防御")
mynl=rs("内力")
mymp=rs("门派")
mysf=rs("身份")
mysex=rs("性别")
mytl=rs("体力")
rs.close
if mywg>10000 and types="copper" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你如此历害的武功已不屑来揭铜榜了。');window.close();}</script>"
	Response.End 
end if
if mywg>20000 and types="silver" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你如此历害的武功已不屑来揭银榜了。');window.close();}</script>"
	Response.End 
end if
if mywg<10000 and types="silver" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你如此差劲的武功还想来揭银榜？回去练练吧！');window.close();}</script>"
	Response.End 
end if
if mywg<20000 and types="gold" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你如此差劲的武功还想来揭金榜？回去练练吧！');window.close();}</script>"
	Response.End 
end if
connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set rs1=server.createobject("adodb.recordset")
Set conn1=Server.CreateObject("ADODB.CONNECTION")
conn1.open connstr1
rs1.open "select * from "&types&" where 姓名='"&sjjh_name&"'",conn1
if not rs1.eof then
	if (rs1("id")<=3 and id=3) or (rs1("id")<=2 and id=2) then
		rs1.close
		set rs1=nothing
		conn1.close
		set conn1=nothing
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：你已经取得了这个或比这还要高的功名，不需要再争了！');window.close();}</script>"
		Response.End
	end if
end if
rs1.close
ltsay=""
boxh=""
rs1.open "select * from "&types&" where id="&id,conn1
if myfy>rs1("防御") and mygj+mynl>rs1("攻击")+rs1("内力") then
	wugong=mywg-rs1("武功")
	defence=myfy-rs1("防御")
	force=mygj-rs1("攻击")
	neili=mynl-rs1("内力")
	conn1.execute "update "&types&" set 姓名='"&sjjh_name&"',性别='"&mysex&"',门派='"&mymp&"',身份='"&mysf&"',武功="&mywg&",内力="&mynl&",体力="&mytl&",攻击="&mygj&",防御="&myfy&" where id="&id
	sql="update 用户 set 内力=内力+50,防御=防御+50,体力=体力-1000,allvalue=allvalue+10,银两=银两-10000000 where 姓名='"&sjjh_name&"'"
	set rs=conn.execute(sql)
	boxh="恭喜你，你揭榜成功！"
	ltsay="<font color=#9900ff>恭喜<font color=#ff0000>"& sjjh_name &"</font>揭榜成功，登上了<font color=#ff0000>"& typeses & classes &"</font>的宝座!消耗体力1000点、支付银子5000000。站长奖励:内力50点、防御50点、经验10点！</font>" 
else
	conn.execute "update 用户 set  内力=内力-100,体力=体力-100,allvalue=allvalue-50,银两=银两-500000 where 姓名='"&sjjh_name&"'"
	boxh="对不起！你揭榜失败！结果被打得鼻青脸肿的。" 
	ltsay="<font color=#9900ff>倒霉的</font><font color=#ff0000>"& sjjh_name &"</font>揭"&typeses&"失败，结果被打的鼻青脸肿。内力失去100点、体力下降100点、经验减少50点、银两损失500000！" 
end if
rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
conn.close
set conn=nothing
newer=ltsay
says="<font color=red><b>【官府公告】</b></font>" & newer
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
Response.Write "<script language=JavaScript>{alert('提示："&boxh&"');window.close();}</script>"
%>
