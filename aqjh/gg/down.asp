<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
dim max,nem,news
news=0
if request.querystring("shu")="" or request.querystring("shu")=0 then
  nem=10
else
  nem=CINT(request.querystring("shu"))
end if
if request.querystring("maxleg")="" or request.querystring("maxleg")=0 then
   max=20
else
   max=CINT(request.querystring("maxleg"))
end if
lx=trim(request.querystring("lx"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if lx="" then
	sql="SELECT * FROM gg  order by e DESC"
else
	sql="SELECT * FROM gg where d='" & lx & "' order by e DESC"
end if
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "document.write('<p align=center><font color=red>暂时没有公告。</font></p>');"
else
do while not rs.eof and news<nem
if len(rs("a"))>max then
	response.write "document.write('<img src=images/br_dot.gif> <a href=../gg/show.asp?id=" & rs("id") & " target=blank title="& rs("a") &">" & Left(rs("a"),max) & "...</a><br>');"	
else
	response.write "document.write('<img src=images/br_dot.gif> <a href=../gg/show.asp?id=" & rs("id") & " target=blank title="& rs("a") &">" & rs("a") & "</a><br>');"
end if
news=news+1
rs.movenext
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing 
if news>nem then
Response.Write "document.write('<div align=right><a href=" & chr(34) & "../gg/more.asp" & chr(34) & " target=blank><img border=0 src=" & chr(34) & "images/gd.gif" & chr(34) & " alt=" & chr(34) & "更多公告" & chr(34) & ">更多┈</a></div>');"
end if
%>