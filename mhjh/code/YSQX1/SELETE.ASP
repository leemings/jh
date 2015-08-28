<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=20 background="../chatroom/bg1.gif">
<div align=center>
<b><font color="#FF0000">注意：此处操作错，将影响到数据库整个系统！<br>
建议不明白的站长不要使用，非常危险！</font></b>
<p>
各属性分类调整,此项是同时调整江湖所有玩家的属性(必须是数字，默认为零)<br>如：选择了“被杀”,填写框里填写了1000,就是把江湖所有玩家的被杀次数调为1000次</p>
<form method="POST" action="turn1.asp">
  <div align="center"> 
    <select name="sjcz">
      <option value="被杀" selected>被杀</option>
      <option value="杀人">杀人</option>
      <option value="积分">积分</option>
      <option value="等级">等级</option>
      <option value="银两">银两</option>
      <option value="精力">精力</option>
      <option value="内力">内力</option>
      <option value="体力">体力</option>
      <option value="攻击">攻击</option>
      <option value="防御">防御</option>
      <option value="资质">资质</option>
      <option value="道德">道德</option>
      <option value="存款">存款</option>
      <option value="精力">精力</option>
      </select>
    <input type="text" name="shuzhi" value=0 size="12" maxlength="8">         
    <input type="submit" value="调整" name="B1" class="p9">         
    <input type="reset" name="Submit" value="清空">            
  </div>         
  </form>  
    此项是把江湖所有玩家的师傅或者徒弟或者配偶。。等调为“无”<br>如：选择了“师傅”,填写框里默认为“无”,就是把江湖所有玩家的的师傅都调为“无”<form method="POST" action="turn2.asp">  
  <div align="center">   
    <select name="sjcz2">  
      <option value="师傅" selected>师傅</option>
      <option value="徒弟">徒弟</option>
      <option value="配偶">配偶</option>
      <option value="洞房花烛">洞房</option>
      <option value="结拜兄弟">兄弟</option>
    </select>
    <input type="text" name="wu" value="无" size="12" maxlength="10" class=nb readonly>          
    <input type="submit" value="调整" name="B1" class="p9">          
    <input type="reset" name="Submit" value="清空">             
  </div>          
  </form>   
      此项是调整用户特效的，当用户反复多次装备武器就会使“特效”字段超出范围<br>以致不能再准备武器了,站长可以一次过把不能装备屋企的玩家恢复正常（默认值就可以了）<form method="POST" action="turn3.asp">   
  <div align="center">    
    <select name="sjcz3">   
      <option value="特技" selected>特技</option>
    </select>
    <input type="text" name="teji" value=";" size="12" maxlength="10" class=nb readonly>          
    <input type="submit" value="调整" name="B1" class="p9">          
    <input type="reset" name="Submit" value="清空">             
  </div>          
  </form>   
    </div>            
</body>