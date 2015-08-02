<%Response.Expires=-1
username=session("Ba_jxqy_username")
if username=""  then  Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from ÓÃ»§ where ÐÕÃû='"&username&"'"
rst.open sqlstr,conn	
umoney=rst("ÒøÁ½")
rst.Close
set rst=nothing
conn.close
set conn=nothing
%> 
<head><title></title><LINK rel=stylesheet href='../chatroom/css.css'></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=bgcolor%>" background="<%=bgimage%>" >
<form method="post" action="compete.asp"  target=compfrm>
  <table  border="0" cellspacing="0" cellpadding="0" bgcolor="<%=bgcolor%>" width='95%' align=center>
    <tr> 
      <td width="27">0ºÅ</td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse0" >
        </td>
      <td width="24">1ºÅ </td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse1" >
        </td>
      <td width="24">2ºÅ </td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse2" >
        </td>
      <td width="24">3ºÅ </td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse3" >
        </td>
      <td width="24">4ºÅ</td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse4" >
        </td>
      <td width="24">5ºÅ </td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse5" >
       </td>
      <td width="24">6ºÅ </td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse6" >
        </td>
      <td width="24">7ºÅ</td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse7" >
        </td>
      <td width="24">8ºÅ </td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse8" >
        </td>
      <td width="24">9ºÅ</td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse9" >
        </td>
      <td width="32">10ºÅ</td>
      <td width="34">  
        <input  maxlength=4 size=4  name="horse10" >
        </td>
      <td width="32">11ºÅ </td>
      <td width="7">  
        <input  maxlength=4 size=4  name="horse11" >
        </td>
    </tr>
    <tr><td colspan=12 align=right><input type=submit value=" ÏÂ ×¢ "> <input type=button onclick=javascript:top.window.close(); value=' ¹Ø ±Õ '> </td><td colspan=12><%Response.Write username&"£º<input type=text style='background-color:"&bgcolor&";border:0;color:000000' readonly value='"&umoney&"'>"%></td></tr>
  </table>
</FORM>
</body>