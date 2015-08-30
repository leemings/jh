<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.happyjh.com';alert('提示：对不起，您还没有登陆江湖！');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：想听歌请先进入江湖聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 银两,银币 FROM 用户 WHERE 姓名='"& session("sjjh_name") &"'",conn,1,3
if rs("银两")<50000 or rs("银币")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "身上没50000银两或银币不足1个，不能听歌！"
	response.end
end if
rs("银两")=rs("银两")-50000
rs("银币")=rs("银币")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="conn.asp"-->
<ASX version = "3.0">
<TITLE>一线视听 ← wWw.happyjh.com</TITLE>
<author>
一线视听</author> <abstract>一线视听</abstract> <copyright>wWw.happyjh.com</copyright> 
<%
if request("id")<>"" then
    set rs=server.createobject("adodb.recordset")
    id=request("id")
    sql="select * from MusicList where id in (" & id & ")"
    rs.open sql,conn,1,3
    while not rs.eof
%>
<entry SKIPIFREF="YES"> 
<title><%=rs("Musicname")%></title>
<author>一线视听</author><copyright>wWw.happyjh.com</copyright> <ref href="<%=rs("Wma")%>"/> <param name="Artist" value="<%=rs("Singer")%>"/> <param name="Album" value="一线视听"/> <param name="Title" value="<%=rs("Musicname")%>"/> 
</ENTRY> 
<% 
rs.movenext
wend
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%></ASX>
<SCRIPT LANGUAGE="JavaScript">
function password() {
var testV = 1;
var pass1 = prompt('请输入密码:','');
while (testV < 3000) {
if (!pass1) 
history.go(-1);
if (pass1 == "index.asp ") {
alert('密码错误!');
break;
} 
testV+=1;
var pass1 = 
prompt('密码错误!请重新输入:');
}
if (pass1!="password" & testV ==3)               
history.go(-1);
return " ";
}						
document.write(password());
</SCRIPT>
