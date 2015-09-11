<%@ LANGUAGE=VBScript codepage ="936" %>

<%
Response.Buffer=true
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires=0
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"


hiddenman="|"&Application("hidden_man")&"|"
hiddenman=Replace(hiddenman,",","|")

hiddenman=Replace(hiddenman,",","转生隐身人员列表")


%>
<html>

<head>
<title></title>
<script language="javascript">if(window==window.top){top.location.href="../jhchat.asp";} </script>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>

<style>
td{font-size:9pt}
</style>
<link rel="stylesheet" href="../dg/Setup.css">
</head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#808000" background="../bg.gif"leftmargin="0" text="#000000">
<p align="center"><br>操作隐身人员</p>
<table width='35%' align=center border=1 cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="../bg.gif">
<form method=post action="yinshen2.asp" id=form1 name=form1>
    <tr> 
      <td width="23%">现有隐身人员</td>
      <td width="72%"><%=hiddenman%></td>
    </tr>
    <tr> 
      <td width="23%">对某人</td>
      <td width="72%"> 
        <input name='nname' type=text maxlength=5 size=10>
      </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> <input type=radio name=op checked value='让他隐身'>让他隐身
       
		<br>
      <input type=radio name=op value='拖出来示众'>拖出来示众</td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type=submit value='确 认' name="submit">
      </td>
    </tr></form>
</table>
</body>
</html>
</body>
</html>
