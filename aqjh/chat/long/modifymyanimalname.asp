<%@ LANGUAGE=VBScript%>
<%

if Session("aqjh_name")="" then Response.Redirect "../error.asp?id=440"
my=Session("aqjh_name")
grade=Session("aqjh_grade")
myid=session("aqjh_id")
id=Request.QueryString ("id")
if id="" then Response.end
chatbgcolor=aqjh_chatbgcolor
chatcolor=aqjh_chatcolor
chatimage=aqjh_chatimage
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value='';parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style></style>

<link rel="stylesheet" href="READONLY/style1.css"></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" leftmargin="3">
<div align="center"><br>
  <font size="3"> <span style='font-size:9pt'><font size="5" face="隶书" color="#FFFFFF">神 
  龙 窟</font></span></font><br>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("SELECT * FROM myanimal where rest=0 and id="&id),conn
if not(rs.EOF or rs.BOF) then
%> 
    
  <table cellpadding="3" cellspacing="1" border="0" align="center" width="100%" bgcolor="#FFFFFF">
    <form method="post" action="modifymyanimalnameok.asp" id=form1 name=form1 target=ps>
      <tr align="center"> 
        <td colspan="2"> 原名 
          <input type="text" readonly name="animalname" size="10" maxlength="10" value="<%=rs("animalname")%>">
          <br>
          新名 
          <input type="text" name="newname" size="10" maxlength="10">
          <br>
          <input type="submit" name="Submit" value="提交">
        </td>
      </tr>
      <%
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%> 
    </form>
  </table>
  <div align="left"><font color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;<font size="2">你可以给你的神兽取名字！如[破破]等！<br>
    &nbsp;&nbsp;&nbsp;&nbsp;注意：请不要取一些低级下流的名字！ 否则站长有权去除你的神兽！</font></font></div>
</div>
</body>
</html>
</body>
</html>
