<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
stock=Trim(Request.Form("stock"))
price=Request.Form("price")
num=Request.Form("num")
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
rs.Open "股票",conn
msg="<head><link rel='stylesheet' href='READONLY/STYLE.CSS'></head><body background='../../jhimg/bk_hc3w.gif' oncontextmenu=self.event.returnValue=false topmargin=0 ><p align=center ><font color=0000ff size=4>股票分红</font></p><hr><form action=allocatenow.asp method=post id=form1 name=form1><table width=60% align=center border=3><tr><td>股票名称</td><td><select name='stock' size='10'>"
do while not (rs.EOF or rs.BOF)
msg=msg&"<option value='"&rs("股票名称")&"'>"&rs("股票名称")&"</option>"
rs.MoveNext
loop
rs.Close
set rs=nothing
msg=msg&"</select></td></tr><tr><td>每股派息</td><td><input type=text name='money' value='5' size=3 maxlength=3></td></tr><tr align=center><td colspan=2><input type=submit name=opt value=' 分 红 '> <input type=button onclick=javascript:history.back(); value=' 返 回 ' id=button1 name=button1></td></tr></table></form></body>"
Response.Write msg
%>