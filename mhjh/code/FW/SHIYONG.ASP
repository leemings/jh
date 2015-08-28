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
%>
<!--#include file="data.asp"-->
<%      
          sql="select * from 物品 where 属性='药品' and 所有者='" & name & "' and ID="&id
	  rs=conn.execute(sql)
			if rs("数量")<=0 then 
			sql="delete * from 物品 where id=" & id
			conn.execute(sql)
                        set rs=nothing
                        conn.close
                        set conn=nothing
                        Response.Redirect "xiaowu.asp?room=c"
			else
			ti=rs("体力")
			nl=rs("内力")
sql="select 体力,内力,等级,攻击,体力加,内力加,防御 from 用户 where 姓名='"&name&"'"
	  rs=conn.execute(sql)
dj=rs("等级")
tili=rs("体力")
neili=rs("内力")
tilij=rs("体力加")
neilij=rs("内力加")
if  ti>0 and tili>dj*tilij then
%>
<script language=vbscript>
   MsgBox "您的体力已经超过上限了,再吃加体力的药就发福了。"
   location.href = "xiaowu.asp?room=c"
</script>
<%
response.end
end if
if nl>0 and neili>dj*neilij then
%>
<script language=vbscript>
   MsgBox "您的内力已经超过上限,再吃加内力的药就爆炸了。"
   location.href = "xiaowu.asp?room=c"
</script>
<%
response.end
else
			sql="update 物品 set 数量=数量-1 where id=" & id
			conn.execute(sql)
			sql="update 用户 set 体力=体力+" & ti & ",内力=内力+" & ti & "  where 姓名='" & name & "'"
			conn.execute(sql)
                        set rs=nothing
                        conn.close
                        set conn=nothing
                        Response.Redirect "xiaowu.asp?room=c"
		        end if
		        end if
%>