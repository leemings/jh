<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
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
if InStr(Application("aqjh_mai"),"|")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：数据有问题不可以购买！');}</script>"
	Response.End
end if
zt=split(Application("aqjh_mai"),"|")
if ubound(zt)<>7 then 
	Response.Write "<script language=JavaScript>{alert('提示：数据有问题不可以购买！');}</script>"
	Response.End
end if
'姓名0，价钱1，币制2，类别3，物品名4，数量5，时间6
name=trim(zt(0))
yin=abs(int(clng(zt(1))))
bz=trim(zt(2))
lb=trim(zt(3))
zswupin=trim(zt(4))
wusl=abs(int(clng(zt(5))))
towho=zt(7)
Application.Lock
Application("aqjh_mai")=""
Application.UnLock
if name=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：自己卖的东东也要买？');}</script>"
	Response.End 
end if
if towho<>"大家" and towho<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：此物品是卖给["&towho&"]的，你无权购买！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select "&lb&" from 用户 where 姓名='"&name&"'",conn,3,3
if  mywpsl(rs(lb),zswupin)<wusl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：["&name&"]的数量不够("&wusl&")个,购买不成功！');}</script>"
	Response.End
end if
zstemp=abate(rs(lb),zswupin,wusl)
rs.close
rs.open "select "&bz&","&lb&" from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if rs(bz)<yin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的["&bz&"]少于"&yin&"购买不成功！');}</script>"
	Response.End
end if
zstemp1=add(rs(lb),zswupin,wusl)
conn.execute "update 用户 set "&lb&"='"&zstemp&"',"&bz&"="&bz&"+"&yin&" where  姓名='"&name&"'"
conn.execute "update 用户 set "&lb&"='"&zstemp1&"',"&bz&"="&bz&"-"&yin&" where  姓名='"&aqjh_name&"'"
mai="<font color=red><b>【购买消息】</b></font>[##]向{%%}购买物品["&zswupin&"]"&wusl&"个成功，花费"&bz&"："&yin&"！"
rs.close
set rs=nothing
conn.close
set conn=nothing
zj="<a href=javascript:parent.sw('[" & aqjh_name & "]'); target=f2>"& aqjh_name & "</a>"
br="<a href=javascript:parent.sw('[" & name & "]'); target=f2>" & name & "</a>"
mai=Replace(mai,"##",zj,1,3,1)
mai=Replace(mai,"%%",br,1,3,1)
says=mai
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
