<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
chatimage=Application("aqjh_chatimage")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
%><head>
<style type="text/css">
td           { font-family: 宋体; font-size: 12px }
body         { font-family: 宋体; font-size: 12px;}
</style>
</head>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<link rel="stylesheet" href="css.css"><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor='#006699' background="<%=chatimage%>" leftmargin="0" topmargin="0" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="<%=chatimage%>">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="red" size="2"><strong>快乐赌场</strong></font></div></td>
      </tr>
    </table>
	
<table width="140" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#000000"
bordercolordark="#FFFFFF" bgcolor="#006699" background="<%=chatimage%>">
  <tr>
<td>
	<p align="center"><input type=button value='江湖赛马' onClick="window.open('horse.asp','horse','scrollbars=yes,resizable=yes')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button129"></p>
      <p align="center"> 
        <input type=button value='赌局消息' onClick="javascript:s('/赌博$ 消息')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button129">
      </p>
      <p align="center"> 
        <input type=button value='赌场庄家行为'  style="border-style:double; border-width:1.0; background-color:#669900;color:#FFFFFF" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button22">
        <br>
        <input type=button value='银两赌局作庄' onClick="javascript:s('/赌博$ 银作庄')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
        <input type=button value='存点赌局作庄' onClick="javascript:s('/赌博$ 点作庄')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='法力赌局作庄' onClick="javascript:s('/赌博$ 法作庄')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='轻功赌局作庄' onClick="javascript:s('/赌博$ 轻作庄')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center"> 
        <input type=button value='金币赌局吆喝' onClick="javascript:s('/对赌$10')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='银币赌局吆喝' onClick="javascript:s('/赌银币$10')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='法力赌局吆喝' onClick="javascript:s('/赌法力$100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='豆点赌局吆喝' onClick="javascript:s('/赌豆$100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
      </p>
      <p align="center">
        <input type=button value='赌场散家行为'  style="background-color:669900;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button222">
        <br>
        &nbsp;<br>
        <input type=button value='银两赌局押小' onClick="javascript:s('/下注$ 银&小&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='银两赌局押大' onClick="javascript:s('/下注$ 银&大&10000')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='银两赌局押双' onClick="javascript:s('/下注$ 银&双&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='银两赌局押单' onClick="javascript:s('/下注$ 银&单&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      <p align="center">
        <input type=button value='赌金属性赌局' onClick="javascript:s('/赌金属性$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button124">
        <input type=button value='赌木属性赌局' onClick="javascript:s('/赌木属性$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button123">
      </p>
      <p align="center"> 
        <input type=button value='赌水属性赌局' onClick="javascript:s('/赌水属性$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button125">
        <input type=button value='赌火属性赌局' onClick="javascript:s('/赌火属性$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button126">
      </p>
      <p align="center"> 
        <input type=button value='赌土属性赌局' onClick="javascript:s('/赌土属性$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button127">
        <input type=button value='风采幸运数字' onClick="javascript:s('/幸运$ 幸运数字1-1000')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button128">
      <p align="center">
        <input type=button value='银两赌局押大' onClick="javascript:s('/下注$ 银&大&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='银两赌局押小' onClick="javascript:s('/下注$ 银&小&10000')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='银两赌局押双' onClick="javascript:s('/下注$ 银&双&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='银两赌局押单' onClick="javascript:s('/下注$ 银&单&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='存点赌局押大' onClick="javascript:s('/下注$ 点&大&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='存点赌局押小' onClick="javascript:s('/下注$ 点&小&100')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='存点赌局押双' onClick="javascript:s('/下注$ 点&双&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='存点赌局押单' onClick="javascript:s('/下注$ 点&单&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='法力赌局押大' onClick="javascript:s('/下注$ 法&大&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='法力赌局押小' onClick="javascript:s('/下注$ 法&小&100')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='法力赌局押双' onClick="javascript:s('/下注$ 法&双&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='法力赌局押单' onClick="javascript:s('/下注$ 法&单&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='轻功赌局押大' onClick="javascript:s('/下注$ 轻&大&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='轻功赌局押小' onClick="javascript:s('/下注$ 轻&小&100')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='轻功赌局押双' onClick="javascript:s('/下注$ 轻&双&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='轻功赌局押单' onClick="javascript:s('/下注$ 轻&单&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='双人赌博' style="background-color:red;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button22">
      </p>
      <p align="center"> 
        <input type=button value='石头' onClick="javascript:s('/双人赌博$ 1')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214" width=30>
        <input type=button value='剪子' onClick="javascript:s('/双人赌博$ 2')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214" width=30>
        <input type=button value=' 布 ' onClick="javascript:s('/双人赌博$ 3')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214" width=30>
		<br>
      </p>
</td>
</tr>
</table>
</body>