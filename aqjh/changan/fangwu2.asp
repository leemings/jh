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
sex=aqjh_jhdj
if instr(action,"'") then response.end 
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where ID=" & id
rs=conn.execute(sql)
yin=rs("售价")
huzhu=rs("户主")
blv=rs("伴侣")
zt=rs("状态")
if rs("银两")<0 or rs("伴侣银两")<0 or rs("数量")<0  then
%><script language=vbscript>
                         MsgBox "你已经成功卖出."
                         location.href = "fangwu.asp"
                    </script><%
conn.execute "update 房屋 set 户主='无',伴侣='无',数量=0,银两=0,伴侣银两=0 where ID=" & id
conn.close
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select * from 用户 where 姓名='"&aqjh_name&"'"
set rs=conn.execute(sql)
if my=huzhu  and zt="正常" then
      conn.execute "update 用户 set 银两=银两+" & yin & "*1/5 where 姓名='"&aqjh_name&"'"
            %><!--#include file="data1.asp"--><%
      conn.execute "update 房屋 set 户主='无',伴侣='无',数量=0,银两=0,伴侣银两=0 where ID=" & id
set rs=nothing
conn.close
set conn=nothing
	Response.Redirect "fangwu.asp"
else
		%>
 <script language=vbscript>
 MsgBox "交易不成功，原因：你不是这个房子的主人或是你的房子出了点状况!"
 location.href = "fangwu.asp"
 </script>
<%
		       rs.close
set rs=nothing
conn.close
set conn=nothing
		end if
   rs.close
   end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
