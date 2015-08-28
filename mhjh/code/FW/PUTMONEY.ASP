<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
my=session("yx8_mhjh_username")
if my="" then response.redirect "../error.asp?id=016"
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
%>
<!--#include file="data.asp"-->
<%
sql="select 银两,门派,性别 from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
yin=rs("银两")
sex=rs("性别")
if rs("门派")="十八地狱" then
%>
<script language=vbscript>
 MsgBox "你是鬼魂，不能放钱到金库!"
 history.back(-1)
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
if money>yin then
%>
<script language=vbscript>
     MsgBox "错误！您的身上没有那么多银两。"
     history.back(-1)
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
if money<1 or money>100000000 then
%>
<script language=vbscript>
     MsgBox "错误！您填写的数量不能小于1或者大于10000万。"
     history.back(-1)
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
money=abs(int(money))
     sql="update 用户 set 银两=银两-'" & money & "' where 姓名='" & my & "'"
     conn.execute(sql)
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data1.asp"-->
<%
       if sex="男" then
       sql="update 房屋 set 银两=银两+'" & money & "' where 户主='" & my & "' or 伴侣='" & my & "'"
       conn.execute(sql)
       else
       sql="update 房屋 set 伴侣银两=伴侣银两+'" & money & "' where 户主='" & my & "' or 伴侣='" & my & "'"
       conn.execute(sql)
       end if
       conn.close
       set conn=nothing
       Response.Redirect "xiaowu2.asp"

%>
