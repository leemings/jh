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
              sql="select * from 物品 where 属性='礼品' and 所有者='" & name & "' and ID="&id
	      rs=conn.execute(sql)			
			if rs("数量")<=0 then
			sql="delete * from 物品 where id=" & id
			conn.execute(sql)
                        set rs=nothing
                        conn.close
                        set conn=nothing
                        Response.Redirect "xiaowu1.asp?room=j"
			else
                        dd=rs("个性")
			ml=rs("美貌")
			sql="update 物品 set 数量=数量-1 where id=" & id
			conn.execute(sql)
			sql="update 用户 set 个性=个性+" & dd & " ,美貌=美貌+" & ml & " where 姓名='" & name & "'"
			conn.execute(sql)
                        set rs=nothing
                        conn.close
                        set conn=nothing
                        Response.Redirect "xiaowu1.asp?room=j"
		        end if
%>