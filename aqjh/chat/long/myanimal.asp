<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
nickname=session("aqjh_name")
grade=session("aqjh_grade")
chatbgcolor=aqjh_chatbgcolor
chatcolor=aqjh_chatcolor
chatimage=aqjh_chatimage
if chatbgcolor="" then chatbgcolor="008888"
%>

<html>
<head>
<style></style>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>

<link rel="stylesheet" href="READONLY/style1.css">
<link rel="stylesheet" href="../READONLY/STYLE.CSS" type="text/css">
</head>
<body oncontextmenu=self.event.returnValue=false background="../bg.gif" bgproperties="fixed" leftmargin="3">
<div align="center"><br>
<font size="3" color="red">护身神龙</font>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
sql= "SELECT * FROM myanimal where username='"&nickname&"' and rest=0"
rs.open sql,conn,1,1
if not(rs.EOF or rs.BOF) then

%>
  <table cellpadding="3" cellspacing="1" border="0" align="center" width="100%" bgcolor="#FFFFFF">
    <tr> 
      <td colspan="2" valign="middle"> 
        <div align="center"><b>神龙</b></font></div>
      </td>
    </tr>
<%
do while not (rs.eof or rs.bof)
gong=rs("attack")
outtime=rs("outtime")
%> 
    <tr align="center"> 
      <td colspan="2" height="42"><font color="#DD451A">龙名：<%=rs("animalname")%></font><font color="#009900">(</font></font><font size="2" color="#009900"><b><%=rs("lei")%></b>)</font></font></td>
    </tr>
  
    <tr> 
      <td colspan="2" bgcolor="#669900"> 
        <%if outtime=true then%>
        <div align=center><font size="2">目前被因禁中</font></div>
        <%
      else
          %>
        <div align="center"><a href=breed.asp?id=<%=rs("id")%> title="给你的神龙喂食！"><font size="2" color="#FFFFFF" class="p9">养</font></a> 
          <font size="2"><a href=modifymyanimalname.asp?id=<%=rs("id")%> title="给你的神龙取一个好听的名字"><font color="#FFFFFF">改</font></a> 
          <a href=zengmyanimal.asp?id=<%=rs("id")%> title="将你的神龙送给你的朋友！"><font color="#FFFFFF">送</font></a> 
          <a href="javascript:s('/神龙$ <%=rs("animalname")%>')" title=攻击：<%=gong%>><font color="#FFFFFF">攻</font></a></font></div>
      </td>
    </tr>
    
    <%
    end if

rs.movenext
loop
rs.Close
set rs=nothing
conn.close
set conn=nothing
%> 
  </table>
<table width="100%" border="0" cellspacing="2" cellpadding="2" bgcolor="#FFFFFF">
<tr>
      <td><font size="2" color="#FF9900">神龙可用来攻击敌人，当神龙状态不佳时，神龙将不听使唤！</font></td>
</tr>
</table>
<br>
  <%
else
%>
  <font color="red">您现在没有神龙！</font> 
  <%end if%>
</div>
</body>
</html>
</body>
</html>
