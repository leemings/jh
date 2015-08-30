<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%><html><head><title>图片♀wWw.happyjh.com♀</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+"[IMG]"+list+"[/IMG]";parent.f2.document.af.sytemp.focus();}</script>
<style TYPE="text/css">
td{line-height:110%;font-size:20pt;}
body{ font-family: 宋体; font-size: 12px;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
.p9{line-height:150%;font-size:9pt;}
A{color:FFFFFF;text-decoration:none;}
A:Visited{color:FFFFFF;}
A:Active{color:FFFF00;}
A:Hover{color:Black;}
</style>
</head>
<body bgcolor="#006699" text="#FFFFFF" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<a name="1"></a> 
<table border="0" align="center" cellpadding="0" bordercolor="#CCCCCC" width="136">
  <tr><td><table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#CCCCFF" size="2"><strong>表 
              情 库</strong></font></div></td>
      </tr>
    </table></td>
</tr>
</table>
<table border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr> 
    <td height="15" colspan="2" align="center" bgcolor="#006699"> 
      <div align=center class=p9><a href=javascript:history.go(0)>刷新</a></div></td>
  </tr>
<tr>
    <td width="136" align="center" bgcolor="#006699"> 
      <%For t=213 to 320%>
      <a href=javascript:s("<%=t%>.gif")><img border=0 src="pic/<%=t%>.gif" width=40 high=40></a> 
      <%next%>
    </td>
  </tr></table>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        
    <td width="100%" height="28" bgcolor="#003366"> 
      <div align=center class=p9><a href=pic2.asp>[2]</a> <a href=pic3.asp>[3]</a> 
        <a href=pic4.asp>[4]</a> <a href=pic5.asp>[5]</a> <a href=pic6.asp>[6]</a></div></td>
      </tr>
    </table>
</body></html>