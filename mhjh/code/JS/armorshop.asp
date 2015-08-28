<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
msg=Request.QueryString("msg")
if msg="" then msg="武器"
msgarr=array("头盔","盔甲","武器","防具","饰品","暗器")
msgtmp="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=ffffff bgcolor=FFB442><tr align=center>"
for i=0 to 5
if msg=msgarr(i) then
msgtmp=msgtmp&"<td bgcolor=ffffff><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"类';return true;"" onmouseout=""window.status='';return true;""><font color=000000>"&msgarr(i)&"</font></a></td>"
else
msgtmp=msgtmp&"<td><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"类';return true;"" onmouseout=""window.status='';return true;""><font color=000000>"&msgarr(i)&"</font></a></td>"
end if
next
msgtmp=msgtmp&"</tr></table>"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,名称,攻击,防御,特效,价格 from 商店 where 属性='"&msg&"' order by 价格 "
rst.open sqlstr,conn
msg="<table width='100%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFffff><tr bgcolor=ffffff align=center><td>名称</td><td>攻击</td><td>防御</td><td>特效</td><td>价格</td><td>数量</td><td>操作</td></TR>"
do while not(rst.EOF or rst.BOF)
msg=msg&"<tr><form method=POST action='buy.asp'><td><input type=hidden checked value='"&rst("id")&"' name='id'>"&rst("名称")&"</td><td>"&rst("攻击")&"</td><td>"&rst("防御")&"</td><td>"&rst("特效")&"</td><td>"&rst("价格")&"</td><td><select name='sl'><option value='1' selected>1</option><option value='2'>2</option><option value='3'>3</option><option value='4'>4</option><option value='5'>5</option><option value='6'>6</option><option value='7'>7</option><option value='8'>8</option><option value='9'>9</option></select></td><td><input type='SUBMIT' name='Submit' value='购买'></td></form></tr>"
rst.MoveNext
loop
msg=msg+"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><LINK href="../style.css" rel=stylesheet><title>武器店</title></head>
<body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'>
<div align=center>
<%=msgtmp%>
   
  <table border="0" width="100%" cellspacing="1" cellpadding="0">
    <tr> 
      <td colspan="11"><table width="100%"  border="1">
          <tr> 
            <td><div align="center"><font color="0">每位大侠只能装备5件武器或者防具，多买了也没用哦。但千万记得要把抵抗4种特效的武器装备好，要不你就惨了。</font></div></td>
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
