<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
msg=Request.QueryString("msg")
if msg="" then msg="����"
msgarr=array("ͷ��","����","����","����","��Ʒ","����")
msgtmp="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff bgcolor=FFB442><tr align=center>"
for i=0 to 5
if msg=msgarr(i) then
msgtmp=msgtmp&"<td bgcolor=ffffff><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"��';return true;"" onmouseout=""window.status='';return true;""><font color=000000>"&msgarr(i)&"</font></a></td>"
else
msgtmp=msgtmp&"<td><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"��';return true;"" onmouseout=""window.status='';return true;""><font color=000000>"&msgarr(i)&"</font></a></td>"
end if
next
msgtmp=msgtmp&"</tr></table>"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,����,����,����,��Ч,�۸� from �̵� where ����='"&msg&"' order by �۸� "
rst.open sqlstr,conn
msg="<table width='100%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFffff><tr bgcolor=ffffff align=center><td>����</td><td>����</td><td>����</td><td>��Ч</td><td>�۸�</td><td>����</td><td>����</td></TR>"
do while not(rst.EOF or rst.BOF)
msg=msg&"<tr><form method=POST action='buy.asp'><td><input type=hidden checked value='"&rst("id")&"' name='id'>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("��Ч")&"</td><td>"&rst("�۸�")&"</td><td><select name='sl'><option value='1' selected>1</option><option value='2'>2</option><option value='3'>3</option><option value='4'>4</option><option value='5'>5</option><option value='6'>6</option><option value='7'>7</option><option value='8'>8</option><option value='9'>9</option></select></td><td><input type='SUBMIT' name='Submit' value='����'></td></form></tr>"
rst.MoveNext
loop
msg=msg+"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><LINK href="../style.css" rel=stylesheet><title>������</title></head>
<body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'>
<div align=center>
<%=msgtmp%>
   
  <table border="0" width="100%" cellspacing="1" cellpadding="0">
    <tr> 
      <td colspan="11"><table width="100%"  border="1">
          <tr> 
            <td><div align="center"><font color="0">ÿλ����ֻ��װ��5���������߷��ߣ�������Ҳû��Ŷ����ǧ��ǵ�Ҫ�ѵֿ�4����Ч������װ���ã�Ҫ����Ͳ��ˡ�</font></div></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td width="10%"><img border="0" src="../image/img21.gif"></td>
      <td width="10%"><img border="0" src="../image/img02.gif"></td>
      <td width="10%"><img border="0" src="../image/img03.gif"></td>
      <td width="10%"><img border="0" src="../image/img04.gif"></td>
      <td width="10%"><img border="0" src="../image/img05.gif"></td>
      <td width="10%"><img border="0" src="../image/img06.gif"></td>
      <td width="10%"><img border="0" src="../image/img06_1.gif"></td>
      <td width="10%"><img border="0" src="../image/img16.gif"></td>
      <td width="10%"><img border="0" src="../image/img29.gif"></td>
      <td width="10%"><img border="0" src="../image/img11.gif"></td>
      <td width="10%"></td>
    </tr>
  </table>
<%=msg%></div>

</body>
