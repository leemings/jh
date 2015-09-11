<%
sql="select * from online where ip='"&Request.ServerVariables("REMOTE_ADDR")&"' and username='"&Request.Cookies("username")&"' or sessionid='"&session.sessionid&"' "
rs.Open sql,Conn,1,3
if rs.eof then rs.addnew
rs("act")=""&act&""
rs("acturl")=""&Request.ServerVariables("script_name")&"?"&Request.ServerVariables("Query_String")&""
rs("forumid")=forumid
rs("bbsname")=""&bbsname&""
rs("lasttime")=now
if Request.Cookies("eremite")<>empty then rs("eremite")=Request.Cookies("eremite")
rs("userface")=membercode
rs("ip")=Request.ServerVariables("REMOTE_ADDR")
rs("username")=Request.Cookies("username")
rs("sessionid")=session.sessionid
rs.update
rs.close
conn.execute("delete from [online] where DateDiff("&SqlChar&"s"&SqlChar&",lasttime,"&SqlNowString&")>"&OnlineTime&" ")
onlinemany=conn.execute("Select count(sessionid)from online")(0)
%>