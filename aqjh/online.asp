<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
document.write("<select name=select style=background-color:#0000FF;font-size:9pt;color:#FFFF00>");
<%
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
next 
%>
document.write("<option selected>爱情江湖有<%=onlinenow%>人聊天</option>");
<%
for i=0 to chatroomnum	
	chatroomxx=split(aqjh_roominfo(i),"|")
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
%>
document.write("<option style='color:FFFFFF'>※※<%=chatroomxx(0)%><%=onlinenum%>人※※</option>");
<%
if onlinenum>0 then 
online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
onlinenum=ubound(online)+1
useronlinename=trim(Application("aqjh_useronlinename"&i))&"  "
'为了隐身功能增加
  aqjh_admin=split(Application("aqjh_admin"),",")
  for j=0 to ubound(aqjh_admin) 
    useronlinename=replace(useronlinename,aqjh_admin(j)&"  ","")
  next
useronlinename=replace(trim(useronlinename),"  ","</option><option>")
%>
document.write("<option><%=useronlinename%></option>")
<%
end if
next%>
document.write("</select>"); 
