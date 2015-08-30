<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<5 then
	Response.Write "<script language=JavaScript>{alert('提示：你不是掌门没有资格进来发放');window.close();}</script>"
	Response.End 
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>

<html>
<head>
<title>掌门发放操作♀一线网络→wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#0099CC" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FFFF00" size="2">掌们发放</font></div></td>
      </tr>
    </table>  
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#336699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang1.asp?cz=桶妖&value=5000" target="d" title="释放魏国妖怪只有魏国人可以打!">桶妖</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
         <div align="center"><font size="-1"><a href="fafang1.asp?cz=木妖&value=5000" target="d" title="释放蜀国妖怪只有蜀国人可以打!">木妖</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
       <div align="center"><font size="-1"><a href="fafang1.asp?cz=人妖&value=5000" target="d" title="释放吴国妖怪只有吴国人可以打!">人妖</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm">
        <div align="center"><font color="#FFFFFF" size="-1">只有掌门才可放妖怪喔</font></div>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
