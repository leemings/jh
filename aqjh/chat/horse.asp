<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM g WHERE c='��' and f=true",conn
xzmd=""
do while not rs.bof and not rs.eof
 xzmd=xzmd&"<option >"&rs("a")&" "& rs("b") &"��"& rs("d")/10000 &"��</option>"
rs.movenext
loop
tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true")
dars=tmprs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<link rel="stylesheet" href="READONLY/STYLE.CSS" type="text/css">
<title>������������</title><body bgcolor="<%=chatbgcolor%>">
<form method="post" action="horseok.asp" target="d">
  <table width="102%" border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#FFCC00">
    <tr align="center" bgcolor="#FFCC66"> 
      <td><font face="����" size="4">������������</font></td>
</tr>

<tr align="center"> 
      <td height="3">
        <select name="1111111">
          <option selected>&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;ע&nbsp;��&nbsp;��&nbsp;&nbsp;</option>
         <%=xzmd%>
        </select>
        <br>
        ���ڻ���<%=(10-dars)%>������!</td>
</tr>
<tr align="center"> 
<td> 
<select name="money">
<option value="500000" selected>50��</option>
<option value="1000000">100��</option>
<option value="2000000">200��</option>
<option value="3000000">300��</option>
<option value="4000000">400��</option>
<option value="5000000">500��</option>
</select>
</td>
</tr>
<tr align="center"> 
<td>ѡ�����</td>
</tr>
<tr align="center"> 
<td> 
<input type="radio" name="radiobutton" value="1" checked>
<img src="horse/1.GIF" width="36" height="35"> <br>
<input type="radio" name="radiobutton" value="2">
<img src="horse/2.GIF" width="36" height="35"><br>
<input type="radio" name="radiobutton" value="3">
<img src="horse/3.GIF" width="36" height="35"><br>
<input type="radio" name="radiobutton" value="4">
<img src="horse/4.GIF" width="36" height="35"><br>
<input type="radio" name="radiobutton" value="5">
<img src="horse/5.GIF" width="36" height="35"><br>
      </td>
</tr>
<tr align="center"> 
<td> 
<input type="submit" name="Submit" value="ȷ��">
<input type="button" name="Submit2" value="����" onClick=this.document.location.href='f3.asp'>
</td>
</tr>
<tr align="center"> 
<td>��СͶע��Ϊ<font color="red">50</font>��<br>
        ���Ͷע��Ϊ<font color="red">500</font>��<br>
        ��ע��10���Զ�����</td>
</tr>
</table>
</form>