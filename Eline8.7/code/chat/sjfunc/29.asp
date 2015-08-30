<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'经验♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_jj then
	Response.Write "<script language=JavaScript>{alert('传点券，需要江湖等级["&jhdj_jj&"]级的才可以操作！');}</script>"
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
says="<font color=red>【传送经验】<font color=" & saycolor & ">"+jingyan(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'经验
function jingyan(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
fn1=abs(fn1)
rs.open "select allvalue,grade FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("allvalue")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有那么多的点券无法传送！');}</script>"
	Response.End
end if
if fn1>100 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：江湖规定，上传点券不能超过100');}</script>"
	Response.End
end if
conn.execute "update 用户 set allvalue=allvalue-" & fn1 & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set allvalue=allvalue+" & fn1 & " where 姓名='" & to1 &"'"
jingyan="##把" & fn1 & "的经验传给了%%,而自己的江湖等级降低了，大无畏呀！"
'记录
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','传授经验','"& jingyan & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>