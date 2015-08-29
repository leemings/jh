<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
nowinroom=session("nowinroom")
name=LCase(trim(request.querystring("name")))
sjjh_name=session("sjjh_name")
db=clng(request.querystring("db"))
dbok=clng(request.querystring("sjjh_bingwen"))
if InStr(name,"or")<>0 or InStr(name,"'")<>0 or InStr(name,"`")<>0 or InStr(name,"=")<>0 or InStr(name,"-")<>0 or InStr(name,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if sjjh_name<>trim(Application("sjjh_dubos")) then
	Response.Write "<script language=JavaScript>{alert('你想作什么，人家也没有说和你赌博！');}</script>"
	Response.End
end if
Application.Lock
	Application("sjjh_dubos")=""
Application.unLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
if db-dbok=-2 or db-dbok=1 then
	Response.Write "<script language=JavaScript>{alert('你输了，倒霉催的！哎~~~');}</script>"
	conn.execute "update 用户 set 银两=银两-100000 where 姓名='"&sjjh_name&"'"
	conn.execute "update 用户 set 银两=银两+200000 where 姓名='"&name&"'"
        hunyin="倒霉：["& sjjh_name &"]赌博里输给了{"& name &"}，10万两白花花的银子飞了！"
end if
if db-dbok=-1 or db-dbok=2 then
        Response.Write "<script language=JavaScript>{alert('你胜了，哈哈，10万两哦~~~');}</script>"
	conn.execute "update 用户 set 银两=银两+100000 where 姓名='"&sjjh_name&"'"
	hunyin="恭喜：["& sjjh_name &"]赌博里赢了{"& name &"}10万两银子！"
end if
if db=dbok then
	Response.Write "<script language=JavaScript>{alert('平了，重新来吧！');}</script>"
    conn.execute "update 用户 set 银两=银两+100000 where 姓名='"&name&"'"
	hunyin="哈哈：["& sjjh_name &"]赌博里和{"& name &"}打平了，伟大的头脑不谋而合。"
end if
'rs.close
'set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【双人赌博】</font>"&hunyin		'聊天数据
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
%>
