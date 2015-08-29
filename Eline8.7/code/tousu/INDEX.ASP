<%@ LANGUAGE=VBScript codepage ="936" %>
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
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>

<html>
<head>
<title>提交状纸♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--td {  font-family: 宋体; font-size: 9pt}body {  font-family: 宋体; font-size: 9pt}select {  font-family: 宋体; font-size: 9pt}A {text-decoration: none; font-family: "宋体"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "宋体"; font-size: 9pt} .big {  font-family: 宋体; font-size: 12pt}.txt {  font-family: "宋体"; font-size: 11pt}
--></style>
<script LANGUAGE="javascript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.topic.value=="")
	{
	  alert("状纸题目没有填！")
	  document.FrmAddLink.topic.focus()
	  return false
	 }
else if(document.FrmAddLink.name.value=="")
	{
	  alert("状告何人没有填！")
	  document.FrmAddLink.name.focus()
	  return false
	 }
else if(document.FrmAddLink.play.value=="")
	{
	  alert("对被告的要求没有填！")
	  document.FrmAddLink.play.focus()
	  return false
	 }
else if(document.FrmAddLink.text.value=="")
	{
	  alert("状词没有填！")
	  document.FrmAddLink.text.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body text="#ffffff" link="#ffffff" alink="#ffffff" vlink="#ffffff" leftmargin="0" topmargin="0" bgcolor="#3a4b91">
<table width=590 border=0 cellspacing=0 cellpadding=0 align=center height="19" bgcolor="#0066CC">
  <tr>
    <td height="26" width="86%"><b>提交状纸</b></td>
    <td align=right height="26" width="14%"> 
      <div align="center"><a href=../jh.asp onClick='javascript:history.back()'>返 
        回</a></div>
    </td>
  </tr>
</table>
<form action=newplan.asp method=post  LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
  <table border=0 cellspacing=0 cellpadding=0 align=center width="590">
    <tr bgcolor="#0066CC"> 
      <td height="26" colspan="2"> <b>状纸题目:</b> 
        <input name=topic size=50 maxlength="30">  
        最长30字</td>  
    </tr>  
    <tr>   
      <td height="17" colspan="2"><b>状告何人：   
        <input name=name size=20 maxlength="20">  
        请写完整的姓名，否则不受理</b></td>  
    </tr>  
    <tr bgcolor="#0066CC">   
      <td colspan="2" height="14">   
        <div align="left"><b>要求对被告的处理：  
          <select name="play" size="1">  
            <option value="罚款一万" selected>罚款一万</option>  
            <option value="罚款十万">罚款十万</option>  
            <option value="坐牢">坐 牢</option>  
          </select>  
          罚款全数归受害人所有</b></div>  
      </td>  
    </tr>  
    <tr>   
      <td colspan="2">   
        <div align="center"><b><img src="../jhmp/bbs/image/edit.gif" width="18" height="13">   
          下面写你的状词：<br>  
          </b>   
          <textarea name=text cols=50 rows=15></textarea>  
        </div>  
        <p align="center">   
          <input type=submit value=发表 name="submit" style="background-color:FF0000;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">  
          &nbsp;&nbsp;&nbsp;&nbsp;   
          <input type=reset value=清除 name="reset" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">  
        </p>  
      </td>  
    </tr>  
  </table>      
</form>      
<div align=center>  
  <p>&nbsp;</p><hr size=1 width="80%">  
</div>      
</body>      
</html>     