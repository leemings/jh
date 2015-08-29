<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>站长发放操作♀wWw.51eline.com♀</title>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center"><font color="#FFFF00" face="楷体_GB2312"><%=Application("sjjh_chatroomname")%>站长发放操作</font><br>
  <table width="70%" border="1" cellspacing="3" cellpadding="3" bgcolor="#339933" align="center">
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=银两&value=100000" target="d" title="发放银两，最大值为10万两!">银两</a></font></div>
      </td>
    </tr>
    <tr>
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=金币&value=10" target="d" title="发放金币，最大值为10个!">金币</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=银币&value=10" target="d" title="发放银币，最大值为10个!">银币</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=道德&value=100" target="d" title="发放道德，最大值为100点!">道德</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=魅力&value=100" target="d" title="发放魅力，最大值为100点!">魅力</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=法力&value=100" target="d" title="发放法力，最大值为100点!">法力</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=轻功&value=100" target="d" title="发放轻功，最大值为100点!">轻功</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=金&value=50" target="d" title="发放金属性，最大值为50点!">金属性</a></font></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=木&value=50" target="d" title="发放木属性，最大值为50点!">木属性</a></font></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=水&value=50" target="d" title="发放水属性，最大值为50点!">水属性</a></font></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=火&value=50" target="d" title="发放火属性，最大值为50点!">火属性</a></font></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=土&value=50" target="d" title="发放土属性，最大值为50点!">土属性</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm">  
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=allvalue&value=50" target="d" title="发放存点，最大值为50点!">存点</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=泡豆点数&value=50" target="d" title="发放豆点，最大值为50点!">豆点</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=老虎&value=10" target="d" title="放出老虎，让大家打!">老虎</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=僵尸&value=50000" target="d" title="发放僵尸，最大体力5万点!">僵尸</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=衰哥&value=50000" target="d" title="发放衰哥，最大体力5万点!">衰哥</a></font></div>
      </td>
    </tr>
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=狗狗&value=5000" target="d" title="发放狗狗，最大体力5万点!">狗狗</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=武士&value=50000" target="d" title="发放武士，最大体力5万点!">武士</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=八国联军&value=50000" target="d" title="八国联军入侵中国，最大人数5万人!">八国联军</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=舞女&value=50000" target="d" title="发放舞女，最大体力5万点!">舞女</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=舞男&value=50000" target="d" title="发放舞男，最大体力5万点!">舞男</a></font></div>
      </td>
    </tr>
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=补药&amp;value=1" target="d" title="发放补药!">补药</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=元宝&value=100000" target="d" title="发放元宝，最大银两10万!">元宝</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang1.asp?cz=桶妖&value=1000" target="d" title="释放桶妖!">桶妖</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
         <div align="center"><font size="-1"><a href="../chat/fafang1.asp?cz=木妖&value=1000" target="d" title="释放木妖!">木妖</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
       <div align="center"><font size="-1"><a href="../chat/fafang1.asp?cz=人妖&value=1000" target="d" title="释放人妖!">人妖</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=抢存点&value=50" target="d" title="抢泡点，最大数量50!">抢存点</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=抢豆点&value=50" target="d" title="抢豆点，最大数量50!">抢豆点</a></font></div>
      </td>
    </tr>
    <tr >
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=抢金卡&value=1" target="d" title="抢金卡，最大数量1!">抢金卡</a></font></div>
      </td>
    </tr>
    <tr>
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/sendfy.asp" target="d" title="发放体力，最高发放5千点!">防御</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=体力&value=5000" target="d" title="发放体力，最高发放5千点!">体力</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=内力&value=5000" target="d" title="发放内力，最高发放5千点!">内力</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="../chat/fafang.asp?cz=武功&value=5000" target="d" title="发放武功，最高发放5千点!">武功</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm">
        <div align="center"><font size="-1">选择要发放的操作，点击即可，不要放的太多哟！</font></div>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
