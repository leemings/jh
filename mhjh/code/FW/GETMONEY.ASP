<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
sj=DateDiff("s",Application("yx8_mhjh_dg"),now()) 
if sj<5 then 
s=5-sj 
Response.write "<script language='javascript'>alert ('我是剑侠之神，请您等上["&s&"秒钟]再操作！');setTimeout('history.back();',1000);</script>"
Response.End 
end if 
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock 
my=session("yx8_mhjh_username")
if my="" then
response.redirect "../error.asp?id=016"
response.end
end if
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next


money=clng(request.form("money"))
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"or") or instr(elevalue,"&") or instr(elevalue,"#")then Response.Redirect "../error.asp?id=056"
next
money=Abs(money)
%>
<!--#include file="data.asp"--><%
sql="select 性别 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
sex=rs("性别")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data1.asp"-->
<%
sql="select * from 房屋 where 户主='" & my & "' or  伴侣='" & my & "'"
set rs=conn.execute(sql)
if sex="男" then
yin=rs("银两")
else
yin=rs("伴侣银两")
end if
if money>yin then
%>
<script language=vbscript>
      MsgBox "错误！您的金库里没有那么多银两。"
      location.href = "xiaowu2.asp"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
if money>100000000 or money<1 then
%>
<script language=vbscript>
        MsgBox "错误！您每次最多只能取10000万。"
        location.href = "xiaowu2.asp"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
       if sex="男" then
       sql="update 房屋 set 银两=银两-'" & money & "' where 户主='" & my & "' or 伴侣='" & my & "'"
       conn.execute(sql)
       else
       sql="update 房屋 set 伴侣银两=伴侣银两-'" & money & "' where 户主='" & my & "' or 伴侣='" & my & "'"
       conn.execute(sql)
       end if
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data.asp"-->
<%
       sql="update 用户 set 银两=银两+'" & money & "' where 姓名='" & my & "'"
       conn.execute(sql)
       conn.close
       set conn=nothing
       Response.Redirect "xiaowu2.asp"	
%>