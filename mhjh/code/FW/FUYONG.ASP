<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
lastip=Request.ServerVariables("QUERY_STRING")
if instr(lastip,"or") then response.end
name=session("yx8_mhjh_username")
if name="" then
response.redirect "../error.asp?id=016"
response.end
end if
id=request("id")
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"or") or instr(elevalue,"&") or instr(elevalue,"#")then Response.Redirect "../error.asp?id=056"
next
%>
<!--#include file="data.asp"-->
<%
sql="SELECT 内力,等级,体力 FROM 用户 WHERE 姓名='" & name & "'"
Set Rs=conn.Execute(sql)
if rs("内力")>=rs("等级")*500000 or rs("体力")>=rs("等级")*500000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
   MsgBox "你吃的太多了，小心撑着"
   location.href = "xiaowu1.asp"
</script>
<%
else
         sql="select * from 物品 where 属性='食品' and 所有者='" & name & "' and ID="&id
         set rs=conn.execute(sql)			
         if rs.eof and rs.bof then
%>
<script language="vbscript">
   MsgBox "你家里没有该食品!!!"
   location.href = "xiaowu1.asp"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing	
response.end
end if
		
     if rs("数量")<=0 then 
	sql="delete * from 物品 where ID=" & id
	conn.execute(sql)
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
        Response.Redirect "xiaowu1.asp"
     else
	ti=rs("体力")
	nl=rs("内力")
	sql="update 物品 set 数量=数量-1 where ID=" & id
	conn.execute(sql)
	sql="update 用户 set 体力=体力+" & ti & ",内力=内力+" & nl & "  where 姓名='" & name & "'"
	conn.execute(sql)
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
        Response.Redirect "xiaowu1.asp"
     end if
end if 
%>