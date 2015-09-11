<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<%
if Session("aqjh_name")=""  then response.end
nickname=Session("aqjh_name")
grade=Session("aqjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select 门派,姓名,信箱,性别 from 用户 where 姓名='"&Session("aqjh_name")&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language="vbscript">
  MsgBox "好象江湖没有你这个人啊"
  location.href = "../../exit.asp"
</script>
<%
else
email=rs("信箱")
menpai=rs("门派")
sex=rs("性别")%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<head>
<LINK 
href="../../css.css" rel=stylesheet>
<title><%=title%></title>
</head>
<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">
<form method="post" action="save.asp" align="center">
  <center>
    <table
    border="1" cellspacing="2" width="415" bordercolor="<%=titlelightcolor%>"
    height="310" cellpadding="2">
      <tr>
            
        <td align="center" colspan="2" bgcolor="<%=titledarkcolor%>"><font color="#FFFFFF">【<%=nickname%>】 
          祈求愿望</font></td>
        </tr>
        <tr>
            <td colspan="2" height="23"><p align="center">&nbsp;<font
            color="<%=titlelightcolor%>">带 （*） 必须填写 </font></p>
            </td>
        </tr>
        <tr>
            
        <td align="center" width="80" height="23">
<p
            align="left">您的姓名： </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <input type="text"
            size="15" maxlength="20" name="name" value="<%=rs("姓名")%>"readonly>
          </p>
            </td>
        </tr>

        <tr>
            
        <td align="center" width="80" height="24">
<p
            align="left">您的性别： </p>
            </td>
            
        <td height="24" width="315"> 
          <p align="left">
            <input type="text"
            size="8" maxlength="20" name="sex" value="<%=sex%>"readonly>
          </p>
            </td>
        </tr>
        <tr>
            
        <td align="center" width="80" height="23">
<p
            align="left">您的信箱： </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <input type="text"
            size="15" maxlength="50" name="email" value="<%=email%>"readonly>
          </p>
            </td>
        </tr>
        <tr>
            
        <td align="center" width="80" height="23"> 
          <p
            align="left">您的门派： </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <input type="text"
            size="15" maxlength="50" name="homepage"
            value="<%=menpai%>"readonly>
             </p>
            </td>
        </tr>
        <tr>
            
        <td width="80" height="23">
<p align="left">愿望类别：
            </p>
            </td>
            
        <td height="23" width="315">
<p align="left">
            <select name="wishtype"
            size="1">
              <option value="love">恋　　爱</option>
                <option value="study">学　　业</option>
                <option value="health">健　　康</option>
                <option value="family">家　　庭</option>
                <option value="work">事　　业</option>
                <option value="future">将　　来</option>
                <option value="wealth">财　　富</option>
                <option value="life">生　　活</option>
            </select> （*） </p>
            </td>
        </tr>
        <tr>
            
        <td width="80" class="pt9" height="80">居住地方：</td>
            
        <td class="pt9" width="315" height="80"> 
          <table border="0" width="300">
            <tr> 
              <td width="53"> 
                <input type="radio" name="address" value="陕西" checked>
                陕西</td>
              <td width="70"> 
                <input type="radio" name="address" value="江苏">
                江苏 </td>
              <td width="51"> 
                <input type="radio" name="address" value="河南">
                河南 </td>
              <td width="53"> 
                <input type="radio" name="address" value="河北">
                河北</td>
              <td width="51"> 
                <input type="radio" name="address" value="广东">
                广东 </td>
            </tr>
            <tr> 
              <td width="53" height="19"> 
                <input type="radio" name="address" value="北京">
                北京 </td>
              <td width="70" height="19"> 
                <input type="radio" name="address" value="江西">
                江西</td>
              <td width="51" height="19"> 
                <input type="radio" name="address" value="四川">
                四川 </td>
              <td width="53" height="19"> 
                <input type="radio" name="address" value="福建">
                福建 </td>
              <td width="51" height="19"> 
                <input type="radio" name="address" value="山西">
                山西 </td>
            </tr>
            <tr> 
              <td height="16" width="53"> 
                <input type="radio" name="address" value="上海">
                上海 </td>
              <td height="16" width="70"> 
                <input type="radio" name="address" value="山东">
                山东</td>
              <td height="16" width="51"> 
                <input type="radio" name="address" value="浙江">
                浙江 </td>
              <td height="16" width="53"> 
                <input type="radio" name="address" value="其他">
                其他 </td>
              <td height="16" width="51">　（*） </td>
            </tr>
          </table>
            </td>
        </tr>
        <tr>
            
        <td width="80" height="66">祈求愿望： </td>
            
        <td width="315" height="66"> 
          <textarea name="info" rows="4" cols="35"></textarea>
            （*） </td>
        </tr>
        <tr>
            
        <td colspan="2" bgcolor="<%=titledarkcolor%>" height="15" align=center><font color="#FFFFFF">送出后合起双眼诚心祈求...</font></td>
        </tr>
        <tr>
            
        <td align="center" colspan="2" height="13">
<input
            type="submit" style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" value="送　　出">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" type="reset" value="清　　除"> </td>
        </tr>
    </table>
    </center></div>
</form>
</body>
</html>
<% 
end if
conn.close
set conn=nothing
set rs=nothing%>