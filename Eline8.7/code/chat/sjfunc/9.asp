<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'警告♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_jg then
	Response.Write "<script language=JavaScript>{alert('提示：警告需要管理等级["&grade_jg&"]才可以操作！');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
'call dianzan(towho)
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
says="<font color=green>【警告】<font color=" & saycolor & ">"+jing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'警告
function jing(fn1,to1)
fn1=trim(fn1)
fn1=Replace(fn1,"=","")
fn1=Replace(fn1,"untion","")
fn1=Replace(fn1,chr(39),"")
jing="<font color=red><b>管理员严肃地对%%说:" & fn1 & "! </b></font>"
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & to1 & "','警告',now(),'" & fn1 & "')"
conn.close
set conn=nothing
end function
%>