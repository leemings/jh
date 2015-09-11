<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
my=aqjh_name
if instr(action,"'") then response.end 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 等级,配偶 from 用户 where 姓名='"&aqjh_name&"'",conn
'hy=rs("会员")
dj=rs("等级")
po=rs("配偶")
if dj<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('房源紧张，须等级大于20级以上才可以购买!');location.href = 'fangwu.asp';}</script>"

response.end
end if
if po="无" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你还没有老婆，不能买房屋!');location.href = 'fangwu.asp';}</script>"

response.end
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%><!--#include file="data1.asp"--><%
sql="select * from 房屋 where 户主='" & my & "' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><!--#include file="data1.asp"--><%
sql="select * from 房屋 where ID=" & id
rs=conn.execute(sql)
yin=rs("售价")
huzhu=rs("户主")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银两 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("银两")<=yin and my=huzhu and huzhu<>"无" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('购买不成功，原因：你的银两不够!');location.href = 'fangwu.asp';}</script>"

response.end
end if
     conn.execute "update 用户 set 银两=银两-" & yin & " where 姓名='"&aqjh_name&"'"
          rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
 %><!--#include file="data1.asp"--><%
      sql="update 房屋 set 户主='" & my & "',伴侣='" & po & "' where ID=" & id
	rs=conn.execute(sql)
	conn.close
	Response.Redirect "fangwu.asp"

else %> 
<script language=vbscript>
            MsgBox "您或您的伴侣已经买过房屋了。"
            location.href = "fangwu.asp"
                    </script>
<%
   rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
%>
