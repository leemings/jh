<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<!--#include file="../showchat.asp"-->
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=5){
window.alert("刷东西是吧？喜欢是吗？点啊，刷啊！！");
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
dj=request("dj")
wp_name=request("wp_name")
wp_sl=request("wp_sl")
kk=request("kk")
wp_type=request("wp_type")
'response.write dj&wp_name&wp_sl&Application("npc_wp"&kk)&wp_type
if aqjh_jhdj>dj then
        Response.Write "<script Language=Javascript>alert('提示：你的等级都这么高了，还和小的们去抢物品啊，还要脸吗？');</script>"
	response.end
end if
if aqjh_grade>15 then
        Response.Write "<script Language=Javascript>alert('提示：身为官府人员，就不要再抢物品了吧！');</script>"
	response.end
end if
if Application("npc_wp"&kk)=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有["&wp_name&"]？！');</script>"
	response.end
end if
Application.Lock
Application("npc_wp"&kk)=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
wptemp=add(rs(wp_type),wp_name,wp_sl)
conn.execute "update 用户 set "&wp_type&"='"&wptemp&"' where 姓名='" & aqjh_name &"'"
call showchat("<font color=ff0000>【NPC暴出物】</font><font color=green>["&aqjh_name&"]今天真是走运，捡到NPC被打暴分裂出来的物品<font color=red>["&wp_name&wp_sl&"个]</font></font>")
%>