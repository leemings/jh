<%@ LANGUAGE=VBScript codepage ="936" %>
<%chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%>
<head>
<title>功能区♀wWw.51eline.com♀</title>
<style type="text/css">
<!--
a:active {  color: #0000FF; text-decoration: underline}
a:hover {  color: #FF0000; text-decoration: underline; cursor: hand}
a:link {  color: #0000FF; text-decoration: underline}
a:visited {  color: #0000FF; text-decoration: underline}
-->
</style>
</head>
<body background="bg.gif" oncontextmenu=self.event.returnValue=false leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width='77%' border=0 bordercolorlight=000000 cellspacing=1 cellpadding=1 bordercolor="#CCCCCC" align="center" >
  <tbody> 
  <tr> 
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a  href="#" onClick="window.open('../hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="当要使用药品、装备物品时点这里！"><font size="-1"><img src="img/gn/wupin.gif" width="30" height="19" border="0"></font></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../sf/index.asp" target="_blank"><img src="img/gn/luntan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#" onClick="window.open('../hcjs/jhjs/dan.asp')" ><img src="img/gn/dangpu.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../hcjs/jhjs/yaopu.asp')" ><img src="img/gn/yaodian.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../hcjs/jhjs/wuqi.asp')" ><img src="img/gn/bingqi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../hcjs/card/card.asp"  target="_blank"><img src="img/gn/kapin.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../yamen/laofan.asp')" ><img src="img/gn/laofang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../yamen/minan.asp')" ><img src="img/gn/mingan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" nowrap width="3%"><a href="#"onClick="window.open('../hcjs/yilao.asp')" ><img src="img/gn/liaoshang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a href="work.asp" target="f3"><font size="-1"><img src="img/gn/gongzuo.gif" width="30" height="19" border="0"></font></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a href="#" onClick="window.open('../yamen/regmodi.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')" title="修改自己的江湖注册资料！"><font size="-1"><img src="img/gn/ziliao.gif" width="30" height="19" border="0"></font></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" height="2" width="1%"><a href="../Diary/Diary.asp" target="_blank"><img src="img/gn/rji.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jl/jlmain.asp" target="_blank"><img src="img/gn/dashi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jhjd/jd.asp" target="_blank"><img src="img/gn/jiudian.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../hcjs/xuewu/xuetang.htm" target="_blank"><img src="img/gn/wuguan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jhmp/wuguan.asp" target="_blank"><img src="img/gn/mishi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../top/top.htm" target="_blank"><img src="img/gn/paihang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../gupiao/stock.asp" target="_blank"><img src="img/gn/gupiao.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../dk/Daikuan.asp" target="_blank"><img src="img/gn/daikuan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../bet/betindex.asp" target="_blank"><img src="img/gn/duguan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="12%"><a href="../jhmp/index.asp" target="_blank"><img src="img/gn/menpai.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../dg/dg.asp" target="_blank"><img src="img/gn/dagong.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="58%"><a href="../jhmp/gry.asp" target="_blank"><img src="img/gn/guer.gif" width="30" height="19" border="0"></a></td>
  </tr>
  </tbody> 
</table>
</body>

