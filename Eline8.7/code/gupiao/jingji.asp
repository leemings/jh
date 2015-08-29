<html>
<head>
<title>股票经纪人办公室</title>
<style>
input,body,select,td{font-size:14;line-height:160%}
</style>
</head>
<body bgcolor="#000000">
<p align=center style='font-size:16;color:yellow'>欢迎<%=sjjh_name%>光临股票经纪人的办公室！
<p align=center style='font-size:14;color:blue'><a href=index.asp>返回股票市场</a></p>
<!--#include file="jhb.asp"-->
<%
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sql="select 经纪人 from 客户 where 帐号='"&sjjh_name&"'"
set rs=conn.execute(sql)
if not rs.eof then
jjr=rs(0)
%>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>您的股票经济人<%=jjr%>在此为您服务</p><br>
<a href=cha.asp>查询今日交易</a>
<br>  
<p align=center><a href=index.asp>离开</a></p>
</table></table>  
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
<p align=center style='font-size:16;color:yellow'>申请股票帐户/更换经济人 
<form method=POST action='jingji2.asp' id=form2 name=form2>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
        <table width=100%>
          <tr> 
            <td width="47%">帐号： 
              <input type=text name=name size=10 value="<%=sjjh_name%>" maxlength="0">
            </td>
            <td width="53%">经济人： 
              <select name=jjr size=1>
                <option value="阿Pip">阿Pip</option>
                <option value="千猪格格">千猪格格</option>
                <option value="印度阿三">印度阿三</option>
                <option value="钱夫人">钱夫人</option>
                <option value="阿土伯">阿土伯</option>
                <option value="羊百万">羊百万</option>
                <option value="约翰乔">约翰乔</option>
                <option value="沙隆巴斯">沙隆巴斯</option>
                <option value="忍太郎">忍太郎</option>
                <option value="无钱莲">无钱莲</option>
                <option value="索镙丝">索镙丝</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td colspan=2 align=center> 
              <input type=submit value=确定 id=submit2 name=submit2>
              <input type=reset value="取消" id=reset2 name=reset2>
            </td>
          </tr>
          <tr> 
            <td colspan=2 style='font-size:9pt'> 
              <hr>
              1.立户时，只须填第一个密码；我们将同时为您安排一位股票经纪人，帮助您完成股票交易，每笔买卖将收取交易金额的1%作为酬劳。<br>
              2.修改密码时，第一个密码为旧密码，第二个密码为新密码！由于本系统采用用户加密的办法，不会出现被人偷钱的事情，所以密码可以不必修改！ 
              <br>
              3.立户时帐号可以任意填写，但是每个江湖人士只可以申请一个股票帐号！</td>
          </tr>
            <td width="47%"> 
          </table>
</td></tr>
</table>
</form>
</body>
</html>
