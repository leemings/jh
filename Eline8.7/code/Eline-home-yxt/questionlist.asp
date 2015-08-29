<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('提示：你不是正站长，你不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set Conn=Server.CreateObject("ADODB.Connection") 
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../chat/questions.asp")
dim rootRs
Set rootRs=Server.CreateObject("ADODB.RecordSet")
sql="select * from questlib"
pagesize=50

rootRs.Open sql,conn,1,1%>
<META content="text/html; charset=gb2312" http-equiv=content-type>
<style type='text/css'>");
body{font-family:"黑体";font-size:12pt;}td{font-family:"宋体";font-size:10pt;}
</style>
<body bgcolor="#CCCCCC" background="../jhimg/bk_hc3w.gif" link="#000000" vlink="#000000" alink="#000000">
<table cellspacing=0 cellpadding=0 align=center width=100% height=100%><tr><td valign=center>
<table align=center width=90% cellspacing=0 cellpadding=0 border=1 BORDERCOLORLIGHT=#000000 BORDERCOLORDARK=#ffffff>
 <td colspan=7 class=navy2><center><b>答题库大全</b></td></tr>
 <%If rootRs.Bof OR rootRs.Eof Then%><tr><td colspan=4><center><font color=red>数据库里目前没有准备任何答题。
 <%else
   rootRs.pagesize=pagesize
   page=Request("page")
   if not IsNumeric(page) or page="" then page=1
   if (page-1)<0 then
     page=rootRs.pagecount
   elseif (page-rootRs.pagecount)>0 then
     page=1
   end if
   rootRs.AbsolutePage=page
   RowCount=rootRs.pagesize
   If Not rootRs.Eof Then%>
     <form action="questionlist.asp" method="post">
     <tr><td colspan=7 nowrap><center>
     <font color=990000>第<%=page%>页 共<%=rootRs.pagecount%>页</font>
     <a href="questionlist.asp?page=<%=(page-1)%>">上一页</a>
     <a href="questionlist.asp?page=<%=(page+1)%>">下一页</a>
     <a href='questionmod.asp'>添加</a>
     转到 <input type="text" name="page" size=6 maxlength=20> 页
     <input class=navy1 type="submit" name="cmdGo" value="GO!" title="转到其他页">
     </td></tr></form>
     <tr><td class=navy1 nowrap><center>编号</td><td class=navy1 nowrap><center>类别</td><td class=navy1 nowrap><center>提问</td><td class=navy1 nowrap><center>答案</td><td class=navy1 nowrap><center>奖励</td>
     <td class=navy1 nowrap>修改</td>
     <%Do While Not rootRs.Eof AND RowCount>0%>
       <tr><td><%=rootRs("ID")%></td><td><%=rootRS("类型")%></td><td><%=rootRS("问题")%></td><td><%=rootRS("回答")%></td><td><%=rootRS("银两")%></td>
       <td><a href="questionmod.asp?ID=<%=rootRs("ID")%>">修改</a></td>
       <%rootRs.MoveNext
       RowCount=RowCount-1
     Loop%><tr><td class=navy1 nowrap><center>编号</td><td class=navy1 nowrap><center>类别</td><td class=navy1 nowrap><center>提问</td><td class=navy1 nowrap><center>答案</td><td class=navy1 nowrap><center>奖励</td>
     <td class=navy1 nowrap>修改</td><%End If
   If rootRs.pagecount>1 then%>
     <tr><td colspan=7><center><b>分页：</b>
     <%For i=1 to rootRs.pagecount
       if (i-page)<>0 then%><a href='questionlist.asp?page=<%=i%>'>[<%=i%>] </a>
       <%else%>(<%=i%>) 
       <%end if
     Next
   end if
 End if
 rootRs.close
 Set rootRs=nothing
 conn.close
 Set conn=nothing%> 
</table><br><center>[<a href='questionmod.asp'>添加</a>]  [<a href='javascript:history.go(-1)'>后退</a>]</table></body></html>

