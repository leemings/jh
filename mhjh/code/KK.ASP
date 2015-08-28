<%

Set conn=Server.CreateObject("ADODB.CONNECTION")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sqlstmt = "select 姓名 from 用户 group by 姓名 having count(*)>1"
Set ws = Server.CreateObject("ADODB.Recordset")  
ws.Open sqlstmt,conn,1,1
response.write ws.recordcount&" 个："
for i=1 to ws.recordcount
if ws.eof then
 Exit For 
    else
      user1=ws("姓名")
      sqlstmt = "select 姓名,id from 用户 where 姓名='" & user1 & "'"
      Set rs = Server.CreateObject("ADODB.Recordset")  
      rs.Open sqlstmt,conn,1,1
      for j=1 to rs.recordcount
      if rs.eof then
      Exit For 
    else
      response.write rs("姓名")&" "
      response.write rs("id")&" "
      rs.MoveNext  
      End If  
     Next
  ws.MoveNext  
    End If  
Next
%>
