<%@ LANGUAGE=VBScript%>
<%
if session("aqjh_name")="" or session("aqjh_id")="" then response.end
nickname=session("aqjh_name")
grade=session("aqjh_grade")
id=Request.QueryString ("id")
if id="" then Response.end

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr

chatbgcolor=aqjh_chatbgcolor
chatcolor=aqjh_chatcolor
chatimage=aqjh_chatimage
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value='';parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<link rel="stylesheet" href="../../css.css">
<link rel="stylesheet" href="../READONLY/STYLE.CSS" type="text/css">

<link rel="stylesheet" href="READONLY/style.css"></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" leftmargin="3">
<div align="center"><br>
  <font size="3"> <span style='font-size:9pt'><font size="5" face="隶书" color="#FFFFFF">神 
  龙 窟</font></span></font><br>
  <%
rs.Open ("SELECT * FROM myanimal where rest=0 and id="&id),conn
if not(rs.EOF or rs.BOF) then
animalname=rs("animalname")
%>
</div>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="100%" bgcolor="#FFFFFF">
    <form method="post" action="modifymyanimalnameok.asp" id=form1 name=form1 target=ps>
      <tr align="center"> 
        
      <td colspan="9"> <font size="2">神龙</font> 
        <input type="text" readonly name="animalname" size="10" maxlength="10" value="<%=animalname%>">
        <br>
        <font size="2"> 目前攻击</font> <%=rs("attack")
%> <br>
        <font size="2">最大攻击</font> <%=rs("allattack")
%><br>
      </td>
    </tr>
             <tr align="center" bgcolor="#FFCC33"> 
        
      <td colspan="7" bgcolor="#669900"><font color="#FFFFFF" size="2">龙 粮</font></td>
      <td bgcolor="#669900"><font color="#FFFFFF" size="2">数 量</font></td>
      <td bgcolor="#669900"><font color="#FFFFFF" size="2">喂养</font></td>
      </tr> <%
rs.Close
rs.Open ("select * from 物品 where 拥有者='"&nickname&"' and 类型='龙粮' and 数量>0 "),conn
do while not(rs.BOF or rs.EOF )
%> 

      <tr align="center"> 
        <td colspan="7"><%=rs("物品名")%></td>
        <td> <%=rs("数量")%></td>
        <td><a href=breedok.asp?id=<%=rs("id")%>&animalname=<%=animalname%> target=ps>喂养</a></td>
        <%
rs.MoveNext
loop
%> </tr>
<%
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%> 
    </form>
  </table>
</body>
</html>
</body>
</html>
