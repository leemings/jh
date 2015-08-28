<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if session("yx8_mhjh_inchat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if

if username="" then
%>
<script language="vbscript">
  alert "不能进入，你还没有登录"
  location.href = "../../exit.asp"
</script>
<%
else
sid=Request.QueryString ("sid")
dim uname
uname=username
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("yx8_mhjh_connstr")
sql="select * from 用户 where 姓名='" & uname & "'"
set rs=conn.execute(sql)
yin=rs("银两")
conn.close
set rs=nothing%>
<!--#include file="jhb.asp"-->
<%
sql="select 资金 from 客户 where 帐号='"&uname&"'"
set rs=conn.execute(sql)
   if rs.eof or rs.bof then
%>
<script language="vbscript">
  alert "你还没有股票帐户，不能炒股"
  location.href = "index.asp"
</script>
<%
   else
   nowmoney=rs("资金")
   set rs=nothing
   sql="select * from 股票 where sid="&sid
   set rss=conn.execute(sql)
   if rss("状态")="封" then
   %>
<script language="vbscript">
  alert "今天股市还没开盘或者已经封盘，不能进行交易"
  location.href = "index.asp"
</script>
<%
   else
   
   sql="select 持股数 from 大户 where sid="&sid&" and 帐号='"&uname&"'"
   set rs=conn.execute(sql)
   if rs.eof then 
   sushare=0
   else
   sushare=rs("持股数")
   end if
%>
<script ID="clientEventHandlersJS"
LANGUAGE="javascript">
<!--
function isnumber(c){
 	if ((c>='0') && (c<='9')) 
		return true; 
	else 		
        return false;
     } 

function B1_onclick(){
     if (Number(form1.ushare.value)==""){
      window.alert ("请填写购买数量!"); 
      return false;}
     for (i=0;i<form1.ushare.value.length;i++){
       n=form1.ushare.value.substr(i,1);
         if (!isnumber(n)){
            window.alert ("请输入一个大于0的整数"); 
            return false;}
   }
return true;}
function B3_onclick(){
     if (Number(form2.ushare.value)==""){
      window.alert ("请填写卖出数量!"); 
      return false;}
     for (i=0;i<form2.ushare.value.length;i++){
       n=form2.ushare.value.substr(i,1);
         if (!isnumber(n)){
            window.alert ("请输入一个大于0的整数"); 
            return false;}
   }
return true;}
//-->
</script>
<html>

<head>
<title></title>
</head>

<body bgcolor="#FFFEF4">
<form name="form1" method="post" action="buy.asp">
  <input type="hidden" name="sid" value="<%=sid%>">
  <div align="center">
    <center>
      <p><font color="#006633">股票代号：<%=sid%> 股票数量：<%=rss("流通股票")%> 当前价格：<%=formatnumber(rss("当前价格"),2)%></font></p>
    </center>
  </div>
  <div align="center">
    <center>
      <p><font color="#006633">您的股票帐户里现有资金：<%=formatnumber(yin,0)%>元,拥有这种股票数：<%=sushare%>,您可以买入该股票：<%=formatnumber((yin)/rss("当前价格")-1,0)%>股 
        </font></p>
    </center>
  </div>
  <div align="center">
    <center>
      <p><font color="#006633">购买数量：
        <input name="ushare"
  size="20" class="smallinput">
        <input type="submit" value="购买" name="B1"
  LANGUAGE="javascript" onclick="return B1_onclick()" class="buttonface">
        <input
  type="reset" value="取消" name="B2" class="buttonface">
        </font></p>
    </center>
  </div>
</form>
<form name="form2" method="post" action="sale.asp">
  <font color="#006633">
  <input type="hidden" name="sid" value="<%=sid%>">
  </font>
  <div align="center">
    <center>
      <p><font color="#006633">卖出数量：
        <input
  name="ushare" size="20" class="smallinput">
        <input type="submit" value="卖出" name="B3"
  LANGUAGE="javascript" onclick="return B3_onclick()" class="buttonface">
        <input
  type="reset" value="取消" name="B2" class="buttonface">
        </font></p>
    </center>
  </div>
</form>
<font color="#006633"><%end if%><%if uname="root" then%><center><a href="zhang.asp">涨</a>————<a href="die.asp">跌</a></center><%end if%></font>
</body>
</html>
<%
end if
end if
conn.close
set conn=nothing
%>