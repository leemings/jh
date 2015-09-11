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
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))

id=trim(request("id"))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
if rs("转生")<5 or rs("道德")<3000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：转生不够五次，或道德低于3000，不能建立组!');location.href = 'javascript:history.go(-1)';}</script>"
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
zid=rs("id")
 rs.close
select case id
case "建立组队"

rs.open "select * from [组队] where 组长="&zid&"",conn,1,1
if not(rs.eof or rs.bof) then
id=rs("id")
zuname=rs("组名")
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你已经建立组了，组id（"&id&"）,组名（"&zuname&"）！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
    conn.execute "insert into 组队 (组名,组长,时间,组员,组员数) values ('"&aqjh_name&"','"&zid&"','"&now()&"','"&aqjh_name&"',1)"
   	conn.execute "update 用户 set 组名='"&aqjh_name&"' where 姓名='" & aqjh_name &"'"

   	says="<font color=red>【组队消息】"&aqjh_name&"建立了组,组名（"&aqjh_name&"），赶快联系他一起参加杀怪行动吧。</font>"
	Response.Write "<script language=JavaScript>{alert('提示：你已经建立组了，组名（"&aqjh_name&"）！');location.href = 'index.asp';}</script>"

   end if
 rs.close
case "删除组队"
rs.open "select * from [组队] where 组长="&zid&"",conn,1,1
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有建立组，不要乱闯！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
   			        conn.execute "update 用户 set 组名='无' where 组名='" &aqjh_name&"'"
   			        sqlstr="delete * from 组队 where 组长="&zid&""
conn.Execute sqlstr

   			        
   			        
   			        
   			says="<font color=red>【删除组队】"&aqjh_name&"删除了组,（"&aqjh_name&"），所有组员解散，找别人组吧。</font>"

	Response.Write "<script language=JavaScript>{alert('提示：删除组成功，所有组员解散！');location.href = 'javascript:history.go(-1)';}</script>"


end if

 rs.close


case "添加组员"
to1=trim(request("to1"))
rs.open "select 组队,组名 FROM 用户 WHERE 姓名='"&to1&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"信息出错，请确认是否有此人');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("组队")=false  then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"没有打开组队状态！禁止加组！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

if rs("组名")<>"" and rs("组名")<>"无" then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"已经有组了，不要重复加入，如果不是你的组，让他先退组！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
rs.close
if  Instr(useronlinename,""&to1&"")=0 then
Response.Write "<script Language=Javascript>alert('“" & to1 & "”不在聊天室中，不能加组。');parent.f2.document.af.towho.options[0].value='大家';parent.f2.document.af.towho.options[0].text='大家';location.href = 'javascript:history.go(-1)';</script>"
Response.end
end if

rs.open "select * from [组队] where 组长="&zid&"",conn,1,1
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有建立组，不要乱闯！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
badword=rs("组员")
	if InStr(badword,""&to1&"")<>0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"已经在组里面了啊！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
	if rs("组员数")>=Application("aqjh_zudui") then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"不能加入,组内人数已经达到系统限制的"&Application("aqjh_zudui")&"人！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if



badword=badword&","&to1

 conn.execute "update 用户 set 组名='"&aqjh_name&"' where 姓名='" & to1 &"'"
   			       
 sqlstr="update 组队 set 组员='"&badword&"',组员数=组员数+1 where 组名='" &aqjh_name&"'"
conn.Execute sqlstr

   			says="<font color=#ff0000>【组队消息】</font> <font color=blue>添加组员"&aqjh_name&"把（"&to1&"）加入了自己组，可以共同进步了</font>"

end if 
rs.close

	Response.Write "<script language=JavaScript>{alert('提示："&to1&"添加成功！');location.href = 'javascript:history.go(-1)';}</script>"

case "删除组员"

to1=trim(request("to1"))
if to1=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('提示：组长不可以删除自己，请删除组操作');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

rs.open "select 组队,组名 FROM 用户 WHERE 姓名='"&to1&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"信息出错，请确认是否有此人');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("组队")=false  then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"没有打开组队状态！禁止加组！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

if rs("组名")<>aqjh_name then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"不是你的组员！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
rs.close

rs.open "select * from [组队] where 组长="&zid&"",conn,1,1
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有建立组，不要乱闯！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
else
zuname=rs("组名")

badword=rs("组员")
	if InStr(badword,to1)=0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示："&to1&"不在组里面了啊！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

	badword=Replace(badword,","&to1&"","")

	
	
	
			conn.Execute "update 用户 set 组名='无' where 姓名='"&to1&"'"
		conn.Execute "update 组队 set 组员='"&badword&"',组员数=组员数-1 where 组名='"&zuname&"'"
			conn.Execute "update 用户 set 魅力=魅力-100 where 姓名='"&to1&"'"
    			says="<font color=#ff0000>【组队消息】</font> <font color=blue>踢出组员"&aqjh_name&"把组员（"&to1&"）踢出了组，（"&to1&"）感觉很没有面子，魅力丢失100。</font>"

	end if
	
		Response.Write "<script language=JavaScript>{alert('提示："&to1&"踢出成功！');location.href = 'javascript:history.go(-1)';}</script>"


rs.close
case else
 Response.Write "<script language=JavaScript>{alert('提示：命令有错，请从正确连接使用江湖功能！');location.href = 'javascript:history.go(-1)';}</script>"
 Response.End

end select
					set rs=nothing
			conn.close
			set conn=nothing
		

call showchat(says)
%>

</body>
</html>
