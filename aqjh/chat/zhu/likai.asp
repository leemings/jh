<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.charset="gb2312"%>
<!--#include file="../../showchat.asp"-->

<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"

id=trim(request("id"))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 组队,组名,id FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：信息出错，请确认是否登陆江湖');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

if rs("组队")=false then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请打开组队状态!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("组队")=""  or rs("组队")="无" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有在组里面,状态出错!');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
zuname=rs("组名")
zid=rs("id")
rs.close
 rs.open "select * FROM 组队 WHERE 组名='"&zuname&"'",conn
if rs.eof or rs.bof then

	conn.Execute "update 用户 set 组名='无' where 姓名='"&aqjh_name&"'"

        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：信息出错，无组信息，状态被初始化');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("组长")=zid then 
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：组长不可以踢除自己，请用删除组操作');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
zuzhang=rs("组长")

badword=rs("组员")
	if InStr(badword,""&aqjh_name&"")=0 then
		conn.Execute "update 用户 set 组名='无' where 姓名='"&aqjh_name&"'"

        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：组内无成员["&aqjh_name&"]，状态被初始化');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

	
	badword=Replace(badword,","&aqjh_name&"","")

	
	
	
			conn.Execute "update 用户 set 组名='无' where 姓名='"&aqjh_name&"'"
		conn.Execute "update 组队 set 组员='"&badword&"',组员数=组员数-1 where 组名='"&zuname&"'"

			conn.Execute "update 用户 set 魅力=魅力-500,allvalue=allvalue-50 where 姓名='"&zuzhang&"'"


    			says="<font color=#ff0000>【组队消息】</font> <font color=blue>离开组"&aqjh_name&"离开组（"&zuname&"），组长感觉很没有面子，魅力丢失500，经验掉了50点。</font>"
rs.close

set rs=nothing
conn.close
set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：操作成功，已经离开组队');location.href = 'javascript:history.go(-1)';}</script>"
	

call showchat(says)

%>

</body>
</html>