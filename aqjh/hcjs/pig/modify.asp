<%
Sub Msg (v)
 Response.Write "<script Language=JavaScript>alert('" & v & "');history.back();</script>"
 Response.End
End Sub
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade>5 then
%>
<script language=javascript>
     history.back()
     alert("提示：身为官府不能参赛！")
</script>
<%
response.end
end if
hytitle=request("hytitle")
hyname=request("hyname")
if hyname="" or hytitle="" then
   Msg "不能为空"
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
sql="select * from 养猪风云榜 where 参赛人='"&aqjh_name&"'"
rs.open sql,connt,1,1
if not rs.EOF then 
%>
<script language=javascript>
     history.back()
     alert("警告：你已经报过名了啊！")
</script>
<%
response.end
else
    rs.close
    rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,1,1
    if rs("银两")<100000 then
%>
<script language=javascript>
     history.back()
     alert("提示：参加比赛，没有10万两是不行的！")
</script>
<%
         response.end
    end if
    conn.Execute "update 用户 set 银两=银两-100000 where 姓名='"&aqjh_name&"'"
    sql="INSERT INTO 养猪风云榜 (参赛人,参赛主题,主题说明,猪赛日期) VALUES ('"&aqjh_name&"','"&hyname&"','"&hytitle&"','"&now()&"')"
    connt.Execute (sql)
%>
<script language=javascript>
     window.location.href ="hfyb.asp"
     alert("提示：恭喜你报名成功，就等着中奖吧！")
</script>
<%
end if
%>