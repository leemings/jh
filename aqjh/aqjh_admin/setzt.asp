<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=439"
t1=trim(Request("t1"))
t2=trim(Request("t2"))
username=Request("name")
sl=request("sl")
if t1="" or username="" or sl="" then
   Response.Write "<script language=javascript>alert('请不要输入空值！');history.back();</script>"
   response.end
end if
if not isnumeric(sl) then
   Response.Write "<script language=javascript>alert('数量请使用数字！');history.back();</script>"
   response.end
end if
if t1="会员金卡" and sl>10000 then
   Response.Write "<script language=javascript>alert('金卡也太多了吧，不得超出1万！');history.back();</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM 用户 where 姓名='"&username&"'"
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要调整的用户不存在！请查看是否正确！');history.back();</script>"
else
 id=rs("id")
 if t2="增加" then
  zt=rs(t1)+sl
 elseif t2="减少" then
  zt=rs(t1)-sl
 else
  Response.Write "<script language=javascript>alert('未知错误！');history.back();</script>"
 end if
 zt=int(zt)
 if zt>100000000 then
   Response.Write "<script language=javascript>alert('警告：也太多了吧，总数都出1亿了！');history.back();</script>"
   response.end
 elseif zt<0 then
   Response.Write "<script language=javascript>alert('警告：不能使用户状态为负！');history.back();</script>"
   response.end
 end if
 conn.execute "update 用户 set "&t1&"="&zt&" where id="&id
 conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','修改："&username&"的["&t1&"]记录！')"
Response.Write "<script language=javascript>alert('提示：修改数据成功！');window.location.href='fine.asp';</script>"
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
