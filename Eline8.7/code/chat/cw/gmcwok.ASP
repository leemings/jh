<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
oldname=trim(Request.Form ("oldname"))
newname=trim(Request.Form ("newname"))
if newname="" or oldname="" or oldname=newname then
	Response.Write "<script language=JavaScript>{alert('提示：宠物名不能为空，也不能与原名相同!');}</script>"
	Response.End 
end if
namelen=0
for i=1 to len(newname)
zh=mid(newname,i,1)
zhasc=asc(zh)
if zhasc<0 then
	namelen=namelen+2
else
	namelen=namelen+1
	if CStr(server.URLEncode(zh))<>CStr(zh) then Response.Redirect "../../error.asp?id=120"
end if
next
if namelen>10 then 
	Response.Write "<script language=JavaScript>{alert('提示：宠物名最多为5个汉字！');}</script>"
	Response.end
end if
badword="站长,管理,管理员,发全,发铨,陈发,一线,E线,一线天,色狼,yixiantian,yxt,eline,射精,奸,死,屎,妈,爹,娘,尻,操,王八,逼,贱,狗,婊,表,靠,叉,插,干,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,处男,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,六扇门,官府"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(newname),bad(i))<>0 then 
		Response.Write "<script language=JavaScript>{alert('提示：宠物名不雅，重新输入！');}</script>"
		Response.end
	end if
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select cw from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
if isnull(rs("cw")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有宠物，请去购买！');</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');</script>"
	response.end
end if
rs.close
rs.open "select i from b where a='"&zt(1)&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：宠物数据出错，请重新购买！');</script>"
	response.end
end if

'名字|类别|生日|体力|攻击|心情|照顾数次|照顾时间
'金龍|金龍|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
temp=newname&"|"&zt(1)&"|"&zt(2)&"|"&zt(3)&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&zt(7)
conn.execute "update 用户 set cw='"&temp&"' where  姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000><b>【宠物改名】</b></font>##把给自己的宠物改名为[<font color=red><b>"&newname&"</b></font>]成功……"
zj="<a href=javascript:parent.sw(\'[" & sjjh_name & "]\'); target=f2>"& sjjh_name & "</a>"
'br="<a href=javascript:parent.sw(\'[" & towho & "]\'); target=f2>" & towho & "</a>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
says=Replace(says,"##",zj,1,3,1)
'says=Replace(says,"%%",br,1,3,1)
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('提示：宠物名更改成["&newname&"]成功!');parent.f3.location.href = 'cw.asp';</script>"
%>
