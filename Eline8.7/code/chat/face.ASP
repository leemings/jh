<%@ LANGUAGE=VBScript codepage ="936" %><%response.expires=0
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 名单头像 from 用户 where 姓名='" & sjjh_name &"'",conn
tx=rs("名单头像")
rs.close
set rs=nothing
conn.close
set conn=nothing
%><html>
<head>
<title>更改头像♀一线网络→wWw.51eline.com♀</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{
font-size:9pt;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff;
}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftmargin="0" topmargin="20"><form method="POST" action="changeface.asp"> 
<div align="center"><font color="#ffff00" size="+1"><b>更改头像</b><br>
  </font><br>
<img id=face src="../ico/<%=tx%>-2.gif" alt=个人形象代表><br><br> 
  <select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
    <%for i=1 to 468%>
    <option value='<%=i%>'><%=i%></option>
    <%next%>
  </select>
  <br>
  <br>
  <input type="submit" value="修改" name="B1" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
  </font></div>
</from>
</body>
</html>