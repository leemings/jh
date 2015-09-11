<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
'if session("myroom")<>"h_拥有者" then
	'Response.Write "<script Language=Javascript>alert('提示：你非此房间主人无全操作！');parent.email.style.visibility='hidden';</script>"
	'Response.End
'end if
%>
<html><title>夫妻钱柜</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center">
<tr>
<td align="center"> <p><img src="IMAGES/money1.gif" width="244" height="191"></p></td>
</tr>
<tr>
<td align="center">只有共同的付出努力才会有更多的收获!</td>
</tr>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT h_夫妻钱柜 FROM house WHERE "& session("myroom") &"='" & aqjh_name &"'",conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：你还没有房屋请去购买！');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
money1=rs("h_夫妻钱柜")
rs.close
rs.open "SELECT 配偶,银两 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,1
'mywife=rs("配偶")
myyin=rs("银两")
rs.close
'rs.open "SELECT 姓名 from [用户] WHERE 姓名='" & mywife & "'",conn,1,1
'if rs.Eof or rs.Bof then
	'rs.close
	'set rs=nothing
	'conn.close
	'set conn=nothing
	'Response.Write "<script Language=Javascript>alert('提示：找不到你的配偶资料!');parent.email.style.visibility='hidden';</script>"
	'response.end
'end if
'mywifezh=rs("姓名")
'rs.close
'rs.open "SELECT h_夫妻钱柜 FROM house WHERE h_拥有者='" & mywifezh &"'",conn,1,1
'if rs.Eof or rs.Bof then
	'rs.Close
	'Set rs = Nothing
	'conn.Close
	'Set conn = Nothing
	'Response.Write "<script Language=Javascript>alert('提示：你的配偶并没有房屋，你还不给他买一个？');parent.email.style.visibility='hidden';</script>"
	'Response.End
'end if
'money2=rs("h_夫妻钱柜")
'rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr><form name="form" method="post" action="money1.asp" onsubmit='return(check());'>
<td align="center">
<table width="96%" border="0">
          <tr> 
            <td align="center"> 我要: 
              <select name="fs" id="fs">
                <option value="1">存钱金库</option>
                <option value="2">取钱金库</option>
              </select> <input name="money" type="text" id="money" size="10" maxlength="10" value="10000">
              两<br>
              现有现金：<%=myyin%>两<br>
              我的钱柜有: 
              <input name="money1" type="text" readonly id="money1" size="10" maxlength="10" value="<%=money1%>">
              两 </td>
          </tr>
          <tr>
            <td align="center">
<input type="submit" name="Submit" value="确定">
            </td>
          </tr>
        </table>
</td>
</form>
</tr>

</table>
<SCRIPT language=JavaScript>
var sl=0;
function check(){
var pattern = /^([0-9])+$/;
var money1=parseFloat(document.form.money1.value);
var money=document.form.money.value;
var fs=parseFloat(document.form.fs.value);
if(pattern.test(money)!=true){alert("提示：请输入数字！");return false;}
money=parseFloat(money);
if (fs==1 && money<10000){alert("提示：存入金库中的钱不可以少于1万两！");return false;}
if (fs==2 && money<1000){alert("提示：取出的钱数不能少于1000两！");return false;}
if (fs==2 && money>money1){alert("提示：你的金库中没有这么多的钱！");return false;}
return true; 
}
</SCRIPT>






