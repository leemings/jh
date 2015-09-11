<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
boyid=cint(request("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if request("w")="del" then
if aqjh_grade<>10 then
 response.write "<script>alert('你无权删除，管理级不够~！');history.back(-1);</script>"
 response.end
end if
sql="delete from gry where id="&boyid
conn.execute(sql)
response.redirect "lingyang.asp"
end if
rs.open "select * from 用户 where 姓名='" &aqjh_name & "'",conn
if rs("银两")<10000000 then
	Response.Write "<script language=JavaScript>{alert('提示：银两没有1千万，孤儿不跟你一起走啊！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("boy")<>"" then
	Response.Write "<script language=JavaScript>{alert('提示：你还有孩子啊！搞什么？');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
peiou=rs("配偶")
if peiou="" or peiou="无" then
	Response.Write "<script language=JavaScript>{alert('提示：你还没结婚呢！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
If DateDiff("d",rs("结婚记念日"),date())<30 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你结婚还不到一个月！');location.href = 'javascript:history.go(-1)';</script>"	
	response.end
end if
rs.close
rs.open "select * from gry where id=" &boyid,conn
if rs.eof then
	Response.Write "<script language=JavaScript>{alert('提示：此孤儿不存在！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
else
 boy=rs("boy")
end if
rs.close
zt=split(boy,"|")
if zt(1)="男" then boysex="images/boy.gif" else boysex="images/girl.gif"
conn.execute "update 用户 set boy='"&boy&"',boysex='"&boysex&"',银两=银两-10000000,道德=道德+50 where 姓名='" & aqjh_name & "'"
conn.execute "update 用户 set boy='"&boy&"',boysex='"&boysex&"' where 姓名='" & peiou & "'"
conn.execute "delete * from gry where id="&boyid
call showchat("<font color=red>【领养小孩】</font>"&aqjh_name&"花了1千万两白银从孤儿养领养了一个小孩，道德上涨50点！")
Response.Write "<script language=JavaScript>{alert('恭喜：领养小孩成功，快去喂养吧！');location.href = 'javascript:history.go(-1)';}</script>"
%>