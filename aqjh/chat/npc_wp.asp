<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<!--#include file="../showchat.asp"-->
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=5){
window.alert("ˢ�����ǰɣ�ϲ�����𣿵㰡��ˢ������");
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
        Response.Write "<script Language=Javascript>alert('��ʾ����ĵȼ�����ô���ˣ�����С����ȥ����Ʒ������Ҫ����');</script>"
	response.end
end if
if aqjh_grade>15 then
        Response.Write "<script Language=Javascript>alert('��ʾ����Ϊ�ٸ���Ա���Ͳ�Ҫ������Ʒ�˰ɣ�');</script>"
	response.end
end if
if Application("npc_wp"&kk)=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ��������["&wp_name&"]����');</script>"
	response.end
end if
Application.Lock
Application("npc_wp"&kk)=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
wptemp=add(rs(wp_type),wp_name,wp_sl)
conn.execute "update �û� set "&wp_type&"='"&wptemp&"' where ����='" & aqjh_name &"'"
call showchat("<font color=ff0000>��NPC�����</font><font color=green>["&aqjh_name&"]�����������ˣ���NPC���򱩷��ѳ�������Ʒ<font color=red>["&wp_name&wp_sl&"��]</font></font>")
%>