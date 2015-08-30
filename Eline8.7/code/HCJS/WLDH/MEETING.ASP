<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title>快乐江湖-武林大会♀wWw.happyjh.com♀</title>
<style type="text/css">
td { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt;
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration: 
               underline }
</style>
</head>
<body bgcolor="#000000" text="#FFFFFF">
<table border="0" height="24" width="650" cellspacing="0" cellpadding="0" align="center">
  <tbody> 
  <tr> 
    <td height="38" width="623"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0"
              bordercolorlight="#000000" bordercolordark="#FFFFFF"
              align="center">
        <tr> 
          <td width="91" height="26"><div align="right"><font size="2">&nbsp; 
                <font
                    color="#FFFFFF"></font><font size="2"><font color="#ffffff"
                    size="2"><span class="zilong"><font color="#FFCC00"> 
                <%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %>
                </font></span></font></font></font></div></td>
          <td width="475" height="26"> 
            <div align="center"> </div>
          </td>
          <td width="104"> 
            <div align="center"> <a href="javascript:window.close()">关闭</a> </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<table width="650" border="0" cellspacing="0" cellpadding="0" align="center" height="298">
  <tr>
    <td width="12" height="305">&nbsp;</td>
    <td width="37" valign="top" height="305"> 
      <div align="center"> </div>
    </td>
    <td valign="top" width="549" height="305"> 
      <div align="center">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="303">
          <tr> 
            <td align="center" height="47"> <img src="images/dtitle2.gif" width="200" height="45"></td>
            <td height="47"> 
              <div align="right"></div>
            </td>
          </tr>
          <tr valign="top" align="center"> 
            <td colspan="2" height="249"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0" height="141">
                <tr> 
                  <td width="11">&nbsp;</td>
                  <td height="13" class="td01">&nbsp;</td>
                  <td width="11">&nbsp;</td>
                </tr>
                <tr> 
                  <td class="td">&nbsp;
                  </td>
                  <td valign="top"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="125">
                      <tr> 
                        <td align="center" height="35" colspan="2"><img src="images/gold.gif" width="120" height="40"></td>
                        <td height="35" align="center" colspan="2"><img src="images/silver.gif" width="120" height="40"></td>
                        <td height="35" align="center" colspan="2"><img src="images/copper.gif" width="120" height="40"></td>
                      </tr>
                      <tr> 
                        <td width="23%" height="22" align="center"> <div align="left">状元：<%
			Set conn=Server.CreateObject("ADODB.CONNECTION")
			Set rs=Server.CreateObject("ADODB.RecordSet")
			conn.open "DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
                        rs.open "select * from gold",conn%>
                            <%=rs("姓名")%> </div></td>
                        <td width="10%" height="22" align="center"><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs.movenext%>
                        </td>
                        <td width="23%" height="22" align="center"><div align="left">状元：<%
                         set rs1=server.createobject("adodb.recordset")
			 rs1.open "select * from silver",conn
			 %>
                            <%=rs1("姓名")%> </div></td>
                        <td width="10%" height="22" align="center"><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs1.movenext%>
                        </td>
                        <td width="23%" height="22" align="center"> <div align="left">状元：<%
                          set rs2=server.createobject("adodb.recordset")
			  rs2.open "select * from copper",conn
			  %>
                            <%=rs2("姓名")%> </div></td>
                        <td width="10%" height="22" align="center"><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs2.movenext%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" height="23" align="center"> <div align="left">榜眼：<%=rs("姓名")%> 
                          </div></td>
                        <td height="23" align="center" width="10%"><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs.movenext%>
                        </td>
                        <td width="23%" height="23" align="center"><div align="left">榜眼：<%=rs1("姓名")%> 
                          </div></td>
                        <td width="10%" height="23" align="center"><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs1.movenext%>
                        </td>
                        <td width="23%" height="23" align="center"> <div align="left">榜眼：<%=rs2("姓名")%> 
                          </div></td>
                        <td width="10%" height="23" align="center"><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a> 
                          <%rs2.movenext%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" height="22" align="center"> <div align="left">探花：<%=rs("姓名")%></div></td>
                        <td height="22" align="center" width="10%"><a href='fight.asp?typename=gold&id=<%=rs("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>
                        <td width="23%" height="22" align="center"><div align="left">探花：<%=rs1("姓名")%></div>
</td>
                        <td width="10%" height="22" align="center"><a href='fight.asp?typename=silver&id=<%=rs1("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>
                        <td width="23%" height="22" align="center"> <div align="left">探花：<%=rs2("姓名")%></div></td>
                        <td width="10%" height="22" align="center"><a href='fight.asp?typename=copper&id=<%=rs2("ID")%>'><img src="images/fight.gif" width="80" height="30" border="0"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td class="td">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="13">&nbsp;</td>
                  <td height="13" class="td01">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
              <br>
              <table width="80%" border="0" cellspacing="0" cellpadding="0" height="49">
                <tr>
                  <td height="60"><font color="#FF0000">江湖提醒</font>： 各位来挑榜的大侠，揭铜榜需武功<font color="#FF0000">300</font>以下，揭银榜需武功<font color="#FF0000">300－800</font>，揭金榜需武功<font color="#FF0000">800</font>以上。在揭榜前请三思而后行，揭榜成功可以将名字留在这里，在聊天室通知大家，并成为后来人的挑战对象，还有一定的奖励。但是挑榜不成功会有一定的损失，具体是什么,大侠一试便知。^_^</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
    </td>
    <td width="25" height="305">&nbsp;</td>
  </tr>
</table>
<div align="center">
</div>
</body>
</html>
<%
rs.close
set rs=nothing
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing
conn.close
set conn=nothing
%>