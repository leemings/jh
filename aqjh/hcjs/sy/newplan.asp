<%if Session("aqjh_name")=""  then
  Response.Redirect "../error.asp?id=440"
else
if request.form("topic")="" then%>
<script language=vbscript>
  MsgBox "请写上状纸的题目！"
  location.href = "javascript:history.back()"
</script>
<%
else
name=request.form("name")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 where 姓名='"&name&"'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
%><script language=vbscript>
 MsgBox "搞错没得，你所告之人不是江湖中人！"
 location.href = "index.asp"
 </script>
<%        
else
name=request.form("name")
topic=request.form("topic")
play=request.form("play")
text=request.form("text")
text=Replace(text,"java","咖啡")
text=Replace(text,"marquee"," ") 
text=Replace(text,"<","小于号") 
text=Replace(text,">","大于号")
text=Replace(text,"&lt"," ")
text=Replace(text,"&gt"," ")
text=Replace(text,"javascript"," ")
text=Replace(text,"\","\\")
text=Replace(text,"\x3c"," ")
text=Replace(text,"\x3e"," ")
text=Replace(text,"\x7d"," ")
text=Replace(text,"/x3c"," ")
text=Replace(text,"/x3e"," ")
text=Replace(text,"/x7d"," ")
text=Replace(text,"\74"," ")
text=Replace(text,"\75"," ")
text=Replace(text,"\76"," ")
text=Replace(text,"\074"," ")
text=Replace(text,"\075"," ")
text=Replace(text,"\076"," ")
text=Replace(text,"/074"," ")
text=Replace(text,"/075"," ")
text=Replace(text,"/076"," ")
text=Replace(text,"oR","或")
text=Replace(text,"or","或")
text=Replace(text,"OR","或")
text=Replace(text,"Or","或")
text=Replace(text,"and","和")
text=Replace(text,"'","无效")
text=Replace(text,"=","无效")
text=Replace(text,"屄","***")
text=Replace(text,"鸡子","***")
text=Replace(text,"婊子","***")
text=Replace(text,"http","无效")
text=Replace(text,"www","无效")
text=Replace(text,"asp","无效")
text=Replace(text,"//","无效")
set rs=server.createobject("adodb.recordset")
sql="select * from sy where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("被告")=name
rs("标题")=topic
rs("要求")=play
rs("状词")=text
rs("原告")=Session("aqjh_name")
rs.update
rs.close
conn.close
set conn=nothing
response.redirect("manage.asp")
		rs.close
		set rs=nothing
end if
end if
end if
%>