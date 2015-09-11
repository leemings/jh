<!-- #include file="setup.asp" -->
<%

id=int(Request("id"))

if Request("menu")="look" then

list=""&Application(CacheName&"vote"&id&"")&""
if list=empty then error2("目前暂时没有记录投票的用户名单！")
error2("参与投票的用户名单：\n"&list&"")

end if


if Request.Cookies("username")=empty then error2("您必须登录后才能投票")

if Request("postvote")="" then error2("请选择，您要投票的项目！")

if instr(Application(CacheName&"vote"&id&""),""&Request.Cookies("username")&" ")>0 or instr(Request.Cookies("vote"),""&id&"|")>0 then error2("您已经投过票了，无需重复投票！")

if instr(Application(CacheName&"voteip"&id&""),""&Request.ServerVariables("REMOTE_ADDR")&" ")>0 then error2("此IP已经投过票了，不能重复投票！")


if Conn.Execute("Select locktopic From [forum] where ID="&id&"")(0)=1 then error("<li>此主题已经关闭，不接受新的投票")




sql="select * from forum where id="&id&""
rs.Open sql,Conn,1,3
for each ho in request.form("postvote")
pollresult=split(rs("pollresult"),"|")
for i = 0 to ubound(pollresult)
if not pollresult(i)="" then
if cint(ho)=i then
pollresult(i)=pollresult(i)+1
end if
allpollresult=""&allpollresult&""&pollresult(i)&"|"
end if
next

rs("pollresult")=allpollresult
rs("lastname")=Request.Cookies("username")
rs("lasttime")=now()
rs.update
allpollresult=""
next


rs.close
Application(CacheName&"vote"&id&"")=""&Request.Cookies("username")&" "&Application(CacheName&"vote"&id&"")&""
Application(CacheName&"voteip"&id&"")=""&Request.ServerVariables("REMOTE_ADDR")&" "&Application(CacheName&"voteip"&id&"")&""
Response.Cookies("vote")=""&Request.Cookies("vote")&""&id&"|"
Response.Cookies("vote").Expires=date+1
error2("投票成功!")


%>