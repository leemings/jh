<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>会员数据库管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理-添加会员卡</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<p align="center">　</p>
<table border="1" width="35%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="141" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="32">
    添加会员卡</td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="99">
    <form method=POST action='addhykok.asp' onsubmit='return(check());' name=addhyk>     
            <p align="left">
      此卡密码：<input type="text" name="kpass" size="20" maxlength="8">&nbsp; 自动生成<input type="checkbox" name="kauto" value="ON" checked>（推荐）</br>
      此卡钱数：<!--webbot bot="Validation" S-Data-Type="Number" S-Number-Separators="x." I-Maximum-Length="6" --><input type="text" name="knum" size="20" maxlength="6"></p>
      <p>
      <input type="submit" value="确定" name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid"></p>
    </form>
    </td>
  </tr>
  </table>
<p align="center">　</p>
<p align="center">本程序由休闲江湖编写</p>
</body>
</html>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="addhyk";
function check()
{
var kpass=document.addhyk.kpass.value;
var knum=document.addhyk.knum.value;
var kauto=document.addhyk.kauto.value;
if(knum=="" || knum==null){window.alert("价格不能没有啊，要不卖谁啊？");return false;}
if(kpass.length<6){window.alert("密码最低要6位！");return false;}
}
//-->
</SCRIPT>

