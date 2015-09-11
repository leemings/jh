<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=50){
window.alert("刷钱是吗？喜欢是吗？点啊，刷啊！！");
i=i+1;}top.location.href="../exit.asp"
}
</script>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
name=LCase(trim(Request.QueryString("name")))
yn=LCase(trim(Request.QueryString("yn")))
if yn=1 and Application("aqjh_dantiao")=aqjh_name  then
'接受单挑
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
myyin=rs("银两")
myzd=rs("内力")+rs("体力")+rs("攻击")+rs("防御")+rs("武功")
rs.close
rs.open "select * FROM 用户 WHERE 姓名='"&name&"'",conn
toyin=rs("银两")
tozd=rs("内力")+rs("体力")+rs("攻击")+rs("防御")+rs("武功")
'我赢了
if myzd>tozd then
	conn.execute "update 用户 set 内力=int(内力/2),银两=银两+"&int(toyin/2)&" where 姓名='" & aqjh_name &"'" &""
	conn.execute "update 用户 set 内力=100,体力=100,武功=100,银两=int(银两/2) where 姓名='"& name &"'"
	dantiao="二人战在一起，战得天昏地暗， {"& name &"} 终于不敌，落荒而逃......；["&aqjh_name&"]得意洋洋得说：这回知道我得厉害了吧！……得到银两："&int(toyin/2)
end if
'我输了
if myzd<tozd then
	conn.execute "update 用户 set 内力=100,体力=100,武功=100,银两=int(银两/2) where 姓名='" & aqjh_name &"'" &""
	conn.execute "update 用户 set 内力=int(内力/2),银两=银两+"&int(myyin/2)&" where 姓名='"& name &"'"
	dantiao="二人战在一起，战得天昏地暗， {"& aqjh_name &"} 终于不敌，落荒而逃......；["&name&"]得意洋洋得说：这回知道我得厉害了吧！……得到银两："&int(myyin/2)
end if
'我们平手
if myzd=tozd then
	conn.execute "update 用户 set 内力=int(内力/2) where 姓名='"& name &"'"
	conn.execute "update 用户 set 内力=int(内力/2) where 姓名='" & aqjh_name &"'"
	dantiao="二人战在一起，远方天昏地暗， 但是由于二人武功不分上下， 互无胜负，只得约定明天再战，草草收队！……" 
end if
dantiao="【单挑决斗】"&aqjh_name&"马鞭一挥高叫道： "& name &" 你个无名小辈！向我挑战，找死！那我就同你大战三百合，让你尝尝某家得厉害！……<br>"&dantiao
else
	if Application("aqjh_dantiao")<>aqjh_name then
		Response.Write "<script language=JavaScript>{alert('提示：网络超时……');}</script>"
		Response.End 
	end if
dantiao="【免战牌】["&aqjh_name&"]环眼一瞪说：哼！ {"& name &"} 尔等小人，也配同我交手？还是拿镜子自己照照吧！……" 
end if
says=dantiao		'聊天数据

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