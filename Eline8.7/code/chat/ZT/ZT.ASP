<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & sjjh_name &"'",conn,3,3
mygj=rs("等级")*sjjh_gjsx+4500 
myfy=rs("等级")*sjjh_fysx+3000
if rs("攻击")>mygj then
	conn.execute("update 用户 set 攻击="& mygj &" where 姓名='"&sjjh_name&"'")
end if
if rs("防御")>myfy then
	conn.execute("update 用户 set 防御="& myfy &" where  姓名='"&sjjh_name&"'")
end if
if rs("会员")=true and clng(DateDiff("d",now(),rs("会员结束")))>0 then
	pdstr="<font size=-1>[泡点制会员]"&rs("会员结束")&"</font>"
else
	pdstr=""
end if
wgj=rs("武功加")
nlj=rs("内力加")
tlj=rs("体力加")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
TD {
	FONT-SIZE: 9pt
}
</style>
<title>♀wWw.happyjh.com♀</title>
</head>

<body topmargin="0" leftmargin="0"  bgcolor="#CCCCCC" style="color: #FFFFFF" oncontextmenu=self.event.returnValue=false >
<table border="0" cellpadding="0" cellspacing="0" width="253">
  <tr>
   <td><img src="img/shim.gif" width="24" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="31" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="45" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="16" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="31" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="19" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="62" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="25" height="1" border="0"></td>
   <td><img src="img/shim.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr valign="top">
   <td colspan="4" align="right"><img name="zt_r02_c1" src="img/zt_r02_c1.jpg" width="116" height="45" border="0"></td>
   <td colspan="4" align="right"><img name="zt_r02_c5" src="img/zt_r02_c5.jpg" width="137" height="45" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="45" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="5" align="right"><img name="zt_r03_c1" src="img/zt_r03_c1.jpg" width="147" height="2" border="0"></td>
   <td rowspan="2" colspan="2" background="img/zt_r03_c6.jpg" align="right"><%=rs("姓名")%></td>
   <td rowspan="28" align="right"><img name="zt_r03_c8" src="img/zt_r03_c8.jpg" width="25" height="301" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="2" border="0"></td>
  </tr>
  <tr valign="top">
   <td rowspan="27" align="right"><img name="zt_r04_c1" src="img/zt_r04_c1.jpg" width="24" height="299" border="0"></td>
   <td rowspan="11" colspan="2" align="right" background="img/zt_r04_c2.jpg"><img src="../../ico/<%=rs("名单头像")%>-2.gif" width="74" height="104"></td>
   <td rowspan="13" colspan="2" align="right"><img name="zt_r04_c4" src="img/zt_r04_c4.jpg" width="47" height="125" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="12" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r05_c6" src="img/zt_r05_c6.jpg" width="81" height="5" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="5" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r06_c6.jpg" align="right"><%=rs("等级")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r07_c6" src="img/zt_r07_c6.jpg" width="81" height="3" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="3" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r08_c6.jpg" align="right"><%=rs("门派")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r09_c6" src="img/zt_r09_c6.jpg" width="81" height="5" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="5" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r10_c6.jpg" align="right"><%=rs("身份")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="16" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r11_c6" src="img/zt_r11_c6.jpg" width="81" height="4" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="4" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r12_c6.jpg" align="right"><%=rs("职业")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r13_c6" src="img/zt_r13_c6.jpg" width="81" height="4" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="4" border="0"></td>
  </tr>
  <tr valign="top">
   <td rowspan="2" colspan="2" background="img/zt_r14_c6.jpg" align="right"><%=rs("配偶")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="12" border="0"></td>
  </tr>
  <tr valign="top">
   <td rowspan="2" colspan="2" align="right"><img name="zt_r15_c2" src="img/zt_r15_c2.jpg" width="76" height="19" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="3" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r16_c6" src="img/zt_r16_c6.jpg" width="81" height="16" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="16" border="0"></td>
  </tr>
  <tr valign="top">
   <td rowspan="14" align="right"><img name="zt_r17_c2" src="img/zt_r17_c2.jpg" width="31" height="174" border="0"></td>
   <td colspan="2" background="img/zt_r17_c3.jpg" align="right"><%=rs("体力")%></td>
   <td rowspan="14" colspan="2" align="right"><img name="zt_r17_c5" src="img/zt_r17_c5.jpg" width="50" height="174" border="0"></td>
   <td background="img/zt_r17_c7.jpg" align="right"><%=rs("道德")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r18_c3" src="img/zt_r18_c3.jpg" width="61" height="5" border="0"></td>
   <td align="right"><img name="zt_r18_c7" src="img/zt_r18_c7.jpg" width="62" height="5" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="5" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r19_c3.jpg" align="right"><%=rs("内力")%></td>
   <td background="img/zt_r19_c7.jpg" align="right"><%=rs("总杀人")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="16" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r20_c3" src="img/zt_r20_c3.jpg" width="61" height="4" border="0"></td>
   <td align="right"><img name="zt_r20_c7" src="img/zt_r20_c7.jpg" width="62" height="4" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="4" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r21_c3.jpg" align="right"><%=rs("武功")%></td>
   <td background="img/zt_r21_c7.jpg" align="right"><%if rs("保护")=true then%>练功保护<%else%>没有保护<%end if%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r22_c3" src="img/zt_r22_c3.jpg" width="61" height="18" border="0"></td>
   <td align="right"><img name="zt_r22_c7" src="img/zt_r22_c7.jpg" width="62" height="18" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="18" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r23_c3.jpg" align="right"><%=rs("allvalue")%></td>
   <td background="img/zt_r23_c7.jpg" align="right"><%=rs("会员等级")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="14" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r24_c3" src="img/zt_r24_c3.jpg" width="61" height="5" border="0"></td>
   <td align="right"><img name="zt_r24_c7" src="img/zt_r24_c7.jpg" width="62" height="5" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="5" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r25_c3.jpg" align="right"><%=rs("攻击")%></td>
   <td background="img/zt_r25_c7.jpg" align="right"><%=rs("会员日期")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="14" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r26_c3" src="img/zt_r26_c3.jpg" width="61" height="5" border="0"></td>
   <td align="right"><img name="zt_r26_c7" src="img/zt_r26_c7.jpg" width="62" height="5" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="5" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r27_c3.jpg" align="right"><%=rs("防御")%></td>
   <td background="img/zt_r27_c7.jpg" align="right"><%=rs("银两")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="14" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r28_c3" src="img/zt_r28_c3.jpg" width="61" height="5" border="0"></td>
   <td align="right"><img name="zt_r28_c7" src="img/zt_r28_c7.jpg" width="62" height="5" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="5" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" background="img/zt_r29_c3.jpg" align="right"><%=rs("知质")%></td>
   <td background="img/zt_r29_c7.jpg" align="right"><%=rs("存款")%></td>
   <td align="right"><img src="img/shim.gif" width="1" height="14" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="2" align="right"><img name="zt_r30_c3" src="img/zt_r30_c3.jpg" width="61" height="30" border="0"></td>
   <td align="right"><img name="zt_r30_c7" src="img/zt_r30_c7.jpg" width="62" height="30" border="0"></td>
   <td align="right"><img src="img/shim.gif" width="1" height="30" border="0"></td>
  </tr>
  <tr valign="top">
   <td colspan="8" align="right"><iframe name=frmhp frameborder=0 width=100% height=100% scrolling="no" src="hp.asp"></iframe></td>
   <td align="right"><img src="img/shim.gif" width="1" height="98" border="0"></td>
  </tr>
</table>
</body>
</html>
<%rs.close
set rs=nothing
conn.close
%>