<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='会员'",conn,2,2
%>
<html>
<head>
<title>世纪网吧加盟</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="img/snowbg.gif">
<div align="center">
  <p><b><font color="#0000CC" size="+2">是的。世纪万维网江湖网吧联盟<br>
    </font></b></p>
  <p align="left">1.网吧加盟的条件？<br>
    a.具有一定规模的网吧。<br>
    <br>
    b.可以提供优质上网服务.<br>
    <br>
    c.在本地区具备一定影响力的网吧.<br>
    <br>
    d.需向世纪万维网江湖提供网吧注册资料！<br>
    <br>
    e.对于IP地址不固定者，需自行修改IP设置。<br>
    <br>
    2.加入收费标准及优惠正策？<br>
    <br>
    a.加盟网吧收费：200元/月(港、奥、台及其它地区以美元结算200$元/月!)<br>
    (<b><font color="#FF0000">春节期间特价:150/月，港、奥、台不变！</font></b>) <br>
    <br>
    b.在加盟网吧上网每天可以多领取工资10万两!<br>
    <br>
    c.在加盟网吧上网，复活后状态不会丢失！<br>
    <br>
    d.在加盟网吧上网，<b><font color="#0000FF">泡点是平常网吧上网的2倍！ </font></b><br>
    <br>
    e.世纪万维网江湖将会在顶部滚动广告支持一周！<br>
    <br>
    f.在加盟网吧上网购物打7折！(会员不变)<br>
    <br>
    <br>
    <br>
    3.加盟办法：<br>
  </p>
  <table width="75%" border="1" align="left">
    <tr>
      <td bgcolor="#0066CC"><%=rs("d")%> 
        <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
      </td>
    </tr>
  </table>
  <p align="left"> <br>
  </p>
</div>
</body>
</html>
