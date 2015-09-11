<%
if session("aqjh_name")="" then response.redirect"../../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if request("id")="" then
sql="select * from QQ order by 用户名 desc"
elseif request("id")=1 then
sql="select * from QQ where 奖励<>0 and 处罚=0"
elseif request("id")=2 then
sql="select * from QQ where 奖励=0 and 处罚<>0"
else
sql="select * from QQ where 奖励=0 and 处罚=0"
end if
rs.open sql,conn,3,2
if rs.eof and rs.bof then str="此系统还没有注册用户!"
if str="" then
	rs.PageSize=20 '每页行数
	pages=rs.pagecount
	records=rs.recordcount
	currentpage=request("currentpage")
	if currentpage="" or currentpage<1 then currentpage=1
	currentpage=cint(currentpage)
	if currentpage>pages then currentpage=pages
	rs.absolutepage=currentpage
else
        currentpage=1
	records=0
	pages=1
end if
%>
<html>
<head>
<title>奖励申请列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF">
<p>◇ <a href="qqlist.asp">申请者列表</a> 
◇ <a href="qqlist.asp?id=1">获奖者名单</a> 
◇ <a href="qqlist.asp?id=2">处罚名单</a> 
◇ <a href="qqlist.asp?id=3">尚未处理名单</a> 
  <a href="send.asp"><b>申请奖励</b></a></p>
<p align="center"><b><font size="3"><%if request("id")="" then%>申请奖励者全部名单<%elseif request("id")=1 then%>获奖者名单<%elseif request("id")=2 then%>处罚名单<%else%>未处理名单<%end if%></font></b></p>
<table width="100%" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000">
  <tr bgcolor="#990000"> 
    <td align="center"><font color="#FFFF00">用户名</font></td>
    <td align="center"><font color="#FFFF00">OICQ号码</font></td>
    <td align="center"><font color="#FFFF00">申请日期</font></td>
    <td align="center"><font color="#FFFF00">是否接受</font></td>
  </tr>    <%linenumber=rs.pagesize%> <%do while (not rs.eof) and (line<linenumber)%> 
  <tr bgcolor="#FFFFFF"> 
    <td><%=rs("用户名")%></td>
    <td><%=rs("oicq")%></td>
    <td align="center"><%=rs("日期")%></td>
    <td><%if rs("奖励")=0 and rs("处罚")=0 then%>正在处理中<%else%><%=rs("奖励说明")%> <%=rs("处罚说明")%><%end if%> 
      <%if session("aqjh_grade")=10 then%>〖<%if rs("奖励")=0 then%><a href="jiangli.asp?uid=<%=rs("用户名")%>&oicq=<%=rs("oicq")%>"> 
      奖励</a><%end if%> <a href="chufa.asp?uid=<%=rs("用户名")%>&oicq=<%=rs("oicq")%>">处罚</a>1 
      <a href="chufa.asp?id=1&uid=<%=rs("用户名")%>&oicq=<%=rs("oicq")%>">处罚2</a>〗〖<a href="delqq.asp?uid=<%=rs("用户名")%>&oicq=<%=rs("oicq")%>">删除</a>〗<%end if%></td>
  </tr><%rs.movenext%> <%line=line+1%> <%loop%> <%rs.close%> 
</table>
<br>
&nbsp;已加入网友共<font color="blue"><%=Records%></font>位&nbsp; 共<font color="blue"><%=Pages%></font>页 
第<font color="red"><%=currentpage%></font>页 <%if currentpage>1 then%> <font size="2"><a href="qqlist.asp?id=<%=request("id")%>&currentpage=<%=currentpage-1%>"><font size="2">[上一页]</font></a> 
<%end if%> <%if currentpage<pages then%> <a href="qqlist.asp?id=<%=request("id")%>&currentpage=<%=currentpage+1%>&search_txt=<%=search_txt%>&px=<%=px%>">[下一页]</a> 
<%end if%> <%if currentpage>1 then%> <a href="qqlist.asp?id=<%=request("id")%>&currentpage=1&search_txt=<%=search_txt%>&px=<%=px%>">[最首页]</a> 
<%end if%> <%if currentpage<pages then%> <a href="qqlist.asp?id=<%=request("id")%>&currentpage=<%=pages%>&search_txt=<%=search_txt%>&px=<%=px%>">[最末页]</a> 
<%end if%> </font> 
</body>
</html>