<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=request("name")
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 攻击,状态 from 用户 where 姓名='" & username & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "你不是江湖中人，不能救人"
conn.close
response.end
else
if rs("攻击")>500000 and (rs("状态")<>"入狱" and rs("状态")<>"死亡" and rs("状态")<>"逮捕") and rs("状态")<>"眠")  then
randomize timer
r=int(rnd*3)
if r=1 then
sql="update 用户 set 状态='正常' where 姓名='" & name & "'"
conn.execute sql
conn.close
Response.Redirect "prison.asp"
else
sql="update 用户 set 状态='入狱' where 姓名='" & username & "'"
conn.execute sql
conn.close
response.write "劫狱不成功！你也被抓了"
response.end
end if
else
response.write "你不能劫狱"
conn.Close             
response.end
end if
end if
%>
