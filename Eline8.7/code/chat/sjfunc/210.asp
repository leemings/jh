<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'请求离派♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_hy then
	Response.Write "<script language=JavaScript>{alert('提示：在线求婚需要["&jhdj_hy&"]级才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【请求离派】<font color=" & saycolor & ">"+qiuhun(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'请求离派
function qiuhun(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 身份<>'掌门' and 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：江湖上找不到你的资料或者你是掌门！');window.close();</script>"
		response.end
end if
if rs("门派")="" or rs("门派")="游侠" or rs("门派")="无" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你并无门派！！');window.close();</script>"
		response.end
end if
if rs("门派")="出家" or rs("门派")="天网"  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('提示：你是出家人，或是天网杀手不可以在这里离开!！！');window.close();</script>"
		response.end
end if
if rs("银两")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你连5万银两离派手续费都没有学什么离派啊？');}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 银两=银两-50000 where 姓名='" & sjjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
qiuhun="[##]向{%%}请求离派：<img src='img/29.gif'>"&fn1&"<input type=button value='同意' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('lipai.asp?name=" & sjjh_name &"&yn=1&to1="&to1&"','d') name=tongyi"&regjm&"><input type=button value='不同意' onClick=javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('lipai.asp?name=" & sjjh_name &"&yn=0&to1="&to1&"','d') name=tongyi"&regjm+1&">"
end function
%>
