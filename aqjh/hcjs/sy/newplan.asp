<%if Session("aqjh_name")=""  then
  Response.Redirect "../error.asp?id=440"
else
if request.form("topic")="" then%>
<script language=vbscript>
  MsgBox "��д��״ֽ����Ŀ��"
  location.href = "javascript:history.back()"
</script>
<%
else
name=request.form("name")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM �û� where ����='"&name&"'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
%><script language=vbscript>
 MsgBox "���û�ã�������֮�˲��ǽ������ˣ�"
 location.href = "index.asp"
 </script>
<%        
else
name=request.form("name")
topic=request.form("topic")
play=request.form("play")
text=request.form("text")
text=Replace(text,"java","����")
text=Replace(text,"marquee"," ") 
text=Replace(text,"<","С�ں�") 
text=Replace(text,">","���ں�")
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
text=Replace(text,"oR","��")
text=Replace(text,"or","��")
text=Replace(text,"OR","��")
text=Replace(text,"Or","��")
text=Replace(text,"and","��")
text=Replace(text,"'","��Ч")
text=Replace(text,"=","��Ч")
text=Replace(text,"��","***")
text=Replace(text,"����","***")
text=Replace(text,"���","***")
text=Replace(text,"http","��Ч")
text=Replace(text,"www","��Ч")
text=Replace(text,"asp","��Ч")
text=Replace(text,"//","��Ч")
set rs=server.createobject("adodb.recordset")
sql="select * from sy where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("����")=name
rs("����")=topic
rs("Ҫ��")=play
rs("״��")=text
rs("ԭ��")=Session("aqjh_name")
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