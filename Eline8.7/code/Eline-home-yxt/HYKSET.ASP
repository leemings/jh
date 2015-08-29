<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_hykmdb")
rs.open "SELECT * FROM sethyk WHERE name='set'",conn
%>
<html>
<head>
<title>会员数据库管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理-高级管理</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<form method="POST" action="hyksetok.asp" name="uphykset" onsubmit='return(check());'>
  <table border="1" width="35%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="209" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="32" colspan="4">
    高级设定</td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="23%" align="center" nowrap height="27">
    泡点会员</td>
    <td bgcolor="#85C2E0"   width="4%" align="center" nowrap height="27">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." s-validation-constraint="Greater than or equal to" s-validation-value="0" --><input type="checkbox" name="phyon" value="ON" <%if rs("phyon") then%>checked<%end if%>></td>
    <td bgcolor="#85C2E0"   width="21%" align="center" nowrap height="27">
    等级会员</td>
    <td bgcolor="#85C2E0"   width="5%" align="center" nowrap height="27">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." s-validation-constraint="Greater than or equal to" s-validation-value="0" --><input type="checkbox" name="dhyon" value="ON" <%if rs("dhyon") then%>checked<%end if%>></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="23%" align="center" nowrap height="28">
    卖银子</td>
    <td bgcolor="#85C2E0"   width="4%" align="center" nowrap height="28">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." s-validation-constraint="Greater than or equal to" s-validation-value="0" --><input type="checkbox" name="mon" value="ON" <%if rs("mon") then%>checked<%end if%>></td>
    <td bgcolor="#85C2E0"   width="21%" align="center" nowrap height="28">
    卖金卡</td>
    <td bgcolor="#85C2E0"   width="5%" align="center" nowrap height="28">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x." s-validation-constraint="Greater than or equal to" s-validation-value="0" --><input type="checkbox" name="jon" value="ON" <%if rs("jon") then%>checked<%end if%>></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="26%" align="center" nowrap height="19" colspan="2">
    　</td>
    <td bgcolor="#85C2E0"   width="25%" align="center" nowrap height="19" colspan="2">
    　</td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    泡点100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="phyf" size="8" value="<%=rs("phyf")%>"></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    泡点一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="phyt" size="8" value="<%=rs("phyt")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    一级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy1f" size="8" value="<%=rs("dhy1f")%>"></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    一级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy1t" size="8" value="<%=rs("dhy1t")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    二级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy2f" size="8" value="<%=rs("dhy2f")%>"></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    二级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy2t" size="8" value="<%=rs("dhy2t")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    三级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy3f" size="8" value="<%=rs("dhy3f")%>"></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    三级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy3t" size="8" value="<%=rs("dhy3t")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    四级100分钟</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy4f" size="8" value="<%=rs("dhy4f")%>"></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    四级一天</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="dhy4t" size="8" value="<%=rs("dhy4t")%>"></td>
  </tr>
  <tr>
    <td bgcolor="#85C2E0"   width="16%" align="center" nowrap height="15">
    1000万银子</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="m" size="8" value="<%=rs("m")%>"></td>
    <td bgcolor="#85C2E0"   width="14%" align="center" nowrap height="15">
    1元金卡</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="15">
    <!--webbot bot="Validation" s-data-type="Number" s-number-separators="x."
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="j" size="8" value="<%=rs("j")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="37" colspan="4">
      <input type="submit" value="确定" name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid"></td>
  </tr>
  </table>
</form>
<p align="center">　</p>
<p align="center">　</p>
<p align="center">本程序由阿男编写</p>
</body>
</html>
<%
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="addhyk";
function check()
{
var phyt=document.uphykset.phyt.value;
var phyf=document.uphykset.phyf.value;
var dhy1t=document.uphykset.dhy1t.value;
var dhy1f=document.uphykset.dhy1f.value;
var dhy2t=document.uphykset.dhy2t.value;
var dhy2f=document.uphykset.dhy2f.value;
var dhy3t=document.uphykset.dhy3t.value;
var dhy3f=document.uphykset.dhy3f.value;
var dhy4t=document.uphykset.dhy4t.value;
var dhy4f=document.uphykset.dhy4f.value;
var m=document.uphykset.m.value;
var j=document.uphykset.j.value;
if(phyt=="" || phyt==null){window.alert("泡点日制会员你没有设定！");return false;}
if(phyf=="" || phyf==null){window.alert("泡点分钟制会员你没有设定！");return false;}
if(dhy1t=="" || dhy1t==null){window.alert("1级等级日制会员你没有设定！");return false;}
if(dhy1f=="" || dhy1f==null){window.alert("1级等级分钟制会员你没有设定！");return false;}
if(dhy2t=="" || dhy2t==null){window.alert("2级等级日制会员你没有设定！");return false;}
if(dhy2f=="" || dhy2f==null){window.alert("2级等级分钟制会员你没有设定！");return false;}
if(dhy3t=="" || dhy3t==null){window.alert("3级等级日制会员你没有设定！");return false;}
if(dhy3f=="" || dhy3f==null){window.alert("3级等级分钟制会员你没有设定！");return false;}
if(dhy4t=="" || dhy4t==null){window.alert("4级等级日制会员你没有设定！");return false;}
if(dhy4f=="" || dhy4f==null){window.alert("4级等级分钟制会员你没有设定！");return false;}
if(m=="" || m==null){window.alert("卖江湖银子比率你没有设定！");return false;}
if(j=="" || j==null){window.alert("卖江湖会员金卡比率你没有设定");return false;}
}
//-->
</SCRIPT>