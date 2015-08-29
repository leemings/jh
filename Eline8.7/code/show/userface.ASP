<!-- #include file="shopconn.asp" -->
<!-- #include file="int.asp" -->
<%
username=trim(request("username"))
if request("sex")="男" then mysex=1 else mysex=0 end if
set rs=server.createobject("adodb.recordset")
sql="select * from [User] where username='"&username&"'"
rs.open sql,connshop,1,1	



if rs.eof then

Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
sqll="SELECT id FROM 用户 WHERE 姓名='"&username&"'"
set rss=conn.execute(sqll)
if  rss.eof then
rss.close
set rss=nothing
conn.close
set conn=nothing
response.write "非注册成员"
response.end

else

set rs1=server.createobject("adodb.recordset")
sql1="select UserID,username from [user] where (userid is null)"
rs1.open sql1,connshop,1,3
rs1.addnew
rs1("UserID")=rss("id")
rs1("username")=username
rs1.update
rs1.close
set rs1=nothing
rss.close
set rss=nothing
conn.close
set conn=nothing
rs.close
set rs=server.createobject("adodb.recordset")
	sql="select userid,visual,period from [user] where username='"&username&"'"
	set rs=connshop.execute(sql)
end if


end if

userid=rs("userid")
myperiod=rs("period")
if isnull(myperiod) or myperiod="" then
	myperiod="||||||||||||||||||||||||"
elseif ubound(split(myperiod,"|"))<>24 then
	myperiod="||||||||||||||||||||||||"
end if
myvisual=GetVisualStr(rs("visual"),mysex,myperiod)

if myvisual<>rs("visual") then connshop.execute("update [user] set visual='"&myvisual&"' where username='"&username&"'")

myvisualsplit=split(myvisual,"|")
myLayerStr=GetLayerStr(myvisualsplit)
mycanshowvisual=isCanShowVisual(myvisualsplit)

rs.close
set rs=nothing
connshop.close
set connshop=nothing
%>
<html>
<head>
</head>
<body oncontextmenu=self.event.returnValue=false >
<table border=0 cellspacing=0 cellpadding=0  align="center">
<tr><td width=140 height=226 id=visualshow ImagePath="<%=PicPath%>" enableLayers=<%=myLayerStr%> usergender=<%=mysex%> baseName=userface/blank.gif <%
						for i=0 to ubound(myvisualsplit)
							if not isnull(myvisualsplit(i)) and trim(myvisualsplit(i))<>"" then
								response.write "layer"&(i+1)&"="&myvisualsplit(i)&" "
							end if
						next
						response.write "style=behavior:url(myshow.htc) setCookies=""document.cookie=this.cookies"" userid="&userid&" localpic="&LocalPic&">"
						%><table width=140 border=0 cellspacing=0 cellpadding=0 ><tr><td height=54 align=center ></td></tr>	</table>
</td></tr></table></body></html>