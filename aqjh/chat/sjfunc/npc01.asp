<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'治愈术
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_roominfo=split(Application("aqjh_room"),";")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【NPC治愈】<bgsound src=wav/FX40A.wav><font color=" & saycolor & ">"+zynpc(mid(says,i+1),towho)+"</font>"
call chatsay("NPC治愈",towhoway,towho,saycolor,addwordcolor,addsays,says)
function zynpc(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
tlj=(rs("等级")*aqjh_tlsx+5260)+rs("体力加")
if rs("银两")<500000 then
   zynpc="##:你没有50万两银子，妙手神医不帮你治疗，现在是金钱社会了！"
   exit function
end if
rs.close
rs.open "select * from npc where n姓名='" & to1 &"'",conn,1,3
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：NPC["&to1&"]并不存在或者它不是NPC你搞错了吧！');}</script>"
	Response.End	
end if
if InStr(";" & Application("aqjh_npc"), ";" &to1& "|")=0 then
   zynpc="提示：NPC[%%]不在线！不能使用治愈术！"
   exit function
end if
if rs("n主人")<>aqjh_name then
  zynpc="##:你不是NPC[%%]的主人，给它治愈，它的主人能放心吗？"
  exit function
end if
n_tl=rs("n体力")
if n_tl<0 then
   n_dengji=int(sqr(int(rs("n经验")/50)))
   rs("n体力")=n_dengji*5000
   rs("n攻击")=n_dengji*100
   rs("n防御")=n_dengji*150
   rs("n内力")=n_dengji*150
   rs.Update
   zynpc= "##对%%使用魔法<b>[治愈术]</b>,因%%接近死亡，还好治愈及时，状态已经全部恢复如初！NPC又充满了往日的杀气，大家可以小心了！"
   exit function
else
  if n_tl>=int(tlj/2) then
    zynpc= "##现在NPC%%的体力大于上限1/2点,虽然使用了治愈术,但是没有效果,白使用了此魔法!"
    Exit Function
 else
    zynpc= "##对%%使用魔法[治愈术],现在%%NPC的体力恢复到上限的1/2点!"
    rs("n体力")=Int(tlj/2)
    rs.Update
    exit function
 end if
end if
end function
%>
