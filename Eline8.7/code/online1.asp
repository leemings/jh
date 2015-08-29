<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
document.write("<select name=select style=background-color:#9ce0ad;font-size:9pt;color:84214d>");
<%
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("sjjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
next 
%>
document.write("<option selected>当前江湖有<%=onlinenow%>人聊天</option>");
<%
for i=0 to chatroomnum	
	chatroomxx=split(sjjh_roominfo(i),"|")
	online=split(trim(Application("sjjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
%>
document.write("<option style='color:blue'>※※<%=chatroomxx(0)%><%=onlinenum%>人※※</option>");
<%
if onlinenum>0 then	
online=split(trim(Application("sjjh_useronlinename"&i)),"  ")
onlinenum=ubound(online)+1
useronlinename=replace(trim(Application("sjjh_useronlinename"&i)),"  ","</option><option>")
%>
document.write("<option><%=useronlinename%></option>")
<%end if
next%>
document.write("</select>");