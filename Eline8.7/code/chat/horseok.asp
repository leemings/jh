<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不能够赌马操作！');}</script>"
	Response.End
end if
money=int(clng(Request.form("money")))
radiobutton=Request.form("radiobutton")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 存款 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("存款")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你钱好像不够！');}</script>"
	Response.End
end if
rs.close
rs.open "select top 1 * FROM g WHERE c='马' and a='"& sjjh_name &"' and f=true",conn,2,2
if not(rs.eof) or not(rs.bof) then
	temp="提示："&sjjh_name&"你买了["&rs("b")&"]号马计："&rs("d")&"两，等着开始吧!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&temp&"');}</script>"
	Response.End 
end if
rs.close
rs.open "select top 1 id,a FROM g WHERE c='马' and f=false",conn,2,2
tempdu=0
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('提示：马赛现在开始了，过一会吧~~!');}</script>"
	tempdu=1
	rs.close
end if
if 	tempdu=0 then
	id=rs("id")
	conn.execute "update 用户 set 存款=存款-"&money&" where 姓名='" & sjjh_name &"'"
	conn.execute "update g set a='"&sjjh_name&"',f=true,d="&money&",b='"&radiobutton&"' where id="&id&" and f=false"
	rs.close
end if
tmprs=conn.execute("Select count(*) As 数量 from g where c='马' and f=true")
dars=tmprs("数量")
'开始赛跑
if dars>=10 then
	randomize(timer())
	mmm = Int(4 * Rnd + 1)
	yinma=""
	rs.open "select * FROM g WHERE c='马' and b='"&mmm&"' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		yinma=yinma&"["&rs("a")&" "& rs("d") &"] "
		conn.execute "update 用户 set 存款=存款+"&rs("d")*3&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		rs.movenext
	loop
	conn.execute "update g set f=false where f=true and c='马'"
	says="<font color=blue><b>[马赛开始]</b></font>紧张激烈的马赛开始了。只见5匹高头竣马快速奔跑………最终<font color=red>"&mmm&"</font>号码<img src=horse/"&mmm&".gif>跑在了最前面，得到了第一名，玩家：<font color=blue><b>"&yinma&"</b></font>玩家幸运押中,每人得到下注的<font color=red>3</font>倍的银两,所有押中玩家内力、体力各加<font color=red>+500</font>！"
	call chat(says)
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End 
end if
set rs=nothing
conn.close
set conn=nothing
says="<font color=blue><b>[江湖马场]</b></font>"&sjjh_name&"在江湖马上选了一匹好马<img src=horse/"&radiobutton&".gif>["&radiobutton&"]号，在这匹马上下注:<font color=red>"&money&"</font>两，现在还差<font color=blue>"&(10-dars)&"</font>开赛！"
call chat(says)
Response.End 

sub chat(says)
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
end sub
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