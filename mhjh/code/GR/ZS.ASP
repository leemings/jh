<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
Response.Redirect "../error.asp?id=016"
else
%>
<HTML>
<HEAD>
<title><%=Application("yx8_mhjh_systemname")%></title>
<link rel="stylesheet" href="../style.css">
</HEAD>
<body  oncontextmenu=self.event.returnValue=false bgcolor="#000000" topmargin="0">

<div align="center">
  <center>
  <table border="0" width="580" cellspacing="0" cellpadding="0" background="../image/bg1.gif" height="420">
    <tr>
      <td width="100%" height="420">
        <p align="center">　<p align="center"><font color="#0000FF" size="3">魂断蓝桥</font></p>
<div align="center">
  <table border="0" width="1%" height="57" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%" height="51">
      <form action="selfmurder.asp" method=post>
<table width="324" border=0 cellspacing="0" cellpadding="0">
<tr><td width="141" height="25">
    <p align="right">姓名：<font color="#FFFF00"><input name=username maxlength=14 size=12 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=username%>" class=nb  readonly></font>
  <center>
    </center>
  <td width="179" height="25">
    <p align="center">
密码：<font color="#FFFF00"><input type=password name=password maxlength=14 size=12 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"></font>
</tr>
  <center>
  </center>
<tr><td align=CENTER width="320" height="27" colspan="2" >
  遗言<input type=text name=say size=30 maxlength=100 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" class=nb  readonly>
  </td></tr> 
<tr><td align=CENTER width="141" height="27" >
  <p align="center">&nbsp;&nbsp; 自杀方式     
  </td><td align=CENTER width="179" height="27" >     
          <select name="select" onChange="javascript:document.forms[0].say.value=this.value;document.forms[0].say.focus();" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" >     
            <option value="由于不小心得罪黑帮,被百来人一阵狂打,用最后的力气咬舌自尽">追杀      
            <option value="被人用麻袋装着，投到了河里，死得悲惨。">落水      
            <option value="拔出一把锈迹斑斑的短剑往脖子上一抹……完了">自刎      
            <option value="服下了几大瓶的安眠药……就次死亡">服毒      
            <option value="关紧房门，解下一条破腰带，然后把自己的脖子和屋梁系在一起，舌头伸的老长。">上吊      
            <option value="然后扑通一声跳到了又脏又臭的河里，死了。">投河      
            <option value="一脚踏空从悬崖上跳了下去，连个全尸也不保！">跳崖      
            <option value="一头撞在又厚又硬的城墙上，然后慢慢的倒了下去!">撞壁      
          </select>  </td></tr>      
<tr><td align=CENTER width="320" height="27" colspan="2" >     
    <p align="center"><input type=submit value=" 自 杀 "  >   &nbsp;</p>         
  </td></tr>         
</table>         
</form>         
</td>         
    </tr>         
  </table>         
</div>         
<p align="center">想清楚才好，自杀后什么都没有了<br>自杀过程中请打开聊天室<br>让大家看到你的悲惨壮举</p>        
</td>    
    </tr>    
  </table>    
  </center>    
</div>    
    
</body></HTML>        
        
<%end if%>