<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="check.asp"-->
<html><head><title>添加NPC</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><link href="jh.css" rel="stylesheet" type="text/css"></head>
<SCRIPT language=JavaScript>
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.form.bds.value;
revisedTitle = addTitle;
document.form.bds.value=revisedTitle;
document.form.bds.focus();
return; }
function check(){
var npcname=document.form.npcname.value;
var npcsex=document.form.npcsex.value;
var npcvalue=document.form.npcvalue.value;
var npcimg=document.form.npcimg.value;
var npcmoney=document.form.npcmoney.value;
var npcgj=document.form.npcgj.value;
var npcfy=document.form.npcfy.value;
var npcwg=document.form.npcwg.value;
var npctl=document.form.npctl.value;
var npcnl=document.form.npcnl.value;
var npcmf=document.form.npcmf.value;
var npcgjl=document.form.npcgjl.value;
var npczh=document.form.npczh.value;
var npccxl=document.form.npccxl.value;
var npcdiren=document.form.npcdiren.value;
var npcqjia=document.form.npcqjia.value;
var npcdie=document.form.npcdie.value;
var npcwp=document.form.npcwp.value;
var npczs=document.form.npczs.value;
var npcccc=document.form.npcccc.value;
var npcwplx=document.form.npcwplx.value;
if(npcname=="" ){alert("提示：NPC的名字不能为空！");return false;}
if(npcsex=="" ){alert("提示：NPc性别不能为空！");return false;}
if(npcimg=="" ){alert("提示：NPc图片不能为空！");return false;}
if(npcdiren=="" ){alert("提示：NPc敌人不能为空！");return false;}
if(npcwp=="" ){alert("提示：NPc物品不能为空！");return false;}
if(npcccc=="" ){alert("提示：NPc出声词不能为空！");return false;}
if(npcwplx=="" ){alert("提示：NPc物品类型不能为空！");return false;}
var pattern = /^([0-9])+$/;
if(pattern.test(npcvalue)!=true){alert("提示：npc经验请使用数字！");return false;}
if(pattern.test(npcmoney)!=true){alert("提示：npc银两请使用数字！");return false;}
if(pattern.test(npcgj)!=true){alert("提示：npc攻击请使用数字！");return false;}
if(pattern.test(npcfy)!=true){alert("提示：npc防御请使用数字！");return false;}
if(pattern.test(npcwg)!=true){alert("提示：npc武功请使用数字！");return false;}
if(pattern.test(npctl)!=true){alert("提示：npc体力请使用数字！");return false;}
if(pattern.test(npcnl)!=true){alert("提示：npc内力请使用数字！");return false;}
if(pattern.test(npcgjl)!=true){alert("提示：npc攻击率请使用数字！");return false;}
if(pattern.test(npccxl)!=true){alert("提示：npc出现率请使用数字！");return false;}
if(pattern.test(npcdie)!=true){alert("提示：npc死次数请使用数字！");return false;}
}
</script>
<form name="form" method="post" action="jhgl.asp?act=addnpc" onSubmit="return check(this);">
  <table width="85%" border=1 align="center" cellpadding=1 cellspacing=0  bordercolor="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor="#d7ebff"> 
      <td colspan="10" align="center">添 加 NPC (<font color="#FF0000"><strong>必须按格式添加,否则不能正常运行</strong></font>)</td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">NPC名字</td>
      <td> <input name="npcname" id="npcname" value="NPC名字" size=10 maxlength=10> 
      </td>
      <td align="center" bgcolor="#d7ebff">性别</td>
      <td> <select name="npcsex" id="npcsex">
          <option value="男" selected>男</option>
          <option value="女">女</option>
        </select></td>
      <td align="center" bgcolor="#d7ebff">经验</td>
      <td> <input name="npcvalue" id="npcvalue" value="0" size=10 maxlength=10> 
      </td>
      <td align="center" bgcolor="#d7ebff">图片</td>
      <td><input name="npcimg" type="text" id="npcimg" value="npc/ma.gif" size="10" maxlength="50"></td>
      <td align="center" bgcolor="#d7ebff">银两</td>
      <td><input name="npcmoney" type="text" id="npcmoney" value="0" size="10" maxlength="15"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">攻击</td>
      <td> <input name="npcgj" id="npcgj" value="0" size=10 maxlength=10> </td>
      <td align="center" bgcolor="#d7ebff">防御</td>
      <td> <input name="npcfy" type="text" id="npcfy" value="0" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">武功</td>
      <td> <input name="npcwg" id="npcwg" value="0" size=10 maxlength=10> </td>
      <td align="center" bgcolor="#d7ebff">体力</td>
      <td><input name="npctl" type="text" id="npctl" value="0" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">内力</td>
      <td><input name="npcnl" type="text" id="npcnl" value="0" size="10" maxlength="10"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">死次数</td>
      <td><input name="npcdie" type="text" id="npcdie" value="0" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">攻击率</td>
      <td><input name="npcgjl" type="text" id="npcgjl" value="0" size="10" maxlength="10"> 
      </td>
      <td align="center" bgcolor="#d7ebff">出现率</td>
      <td><input name="npccxl" type="text" id="npccxl" value="0" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">主人</td>
      <td><input name="npczhuren" type="text" id="npczhuren" value="无" size="10" maxlength="10"></td>
      <td align="center" bgcolor="#d7ebff">敌人</td>
      <td><input name="npcdiren" type="text" id="npcdiren" value="无" size="10" maxlength="10"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">物品</td>
      <td colspan="7"><input name="npcwp" type="text" id="npcwp" value="物品名|数量;物品名2|数量2;" size="60" maxlength="200"></td>
      <td align="center" bgcolor="#d7ebff">物品类型</td>
      <td><input name="npcwplx" type="text" id="npcwplx" value="w1" size="4" maxlength="3"></td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#d7ebff">出场词</td>
      <td colspan="9"><input name="npcccc" type="text" id="npcccc" value="##&lt;img src=npc/ma.gif width=50&gt;进入聊天室!" size="70" maxlength="200"></td>
    </tr>
    <tr> 
      <td colspan="10" align="center" valign=top bgcolor="#d7ebff"> <input type="SUBMIT" name="Submit" value="添加"> 
        <input type="RESET" name="Reset" value="清除"> </td>
    </tr>
  </table>
  <br>
  <table width="85%" border="0" align="center">
    <tr>
      <td><font color="#0000FF"><strong>经验:</strong></font>这个是用处理NPC的等级,他的数值就是:等级x等级x50=经验数,这一个也赶影响到他的攻防值.<br>
        <font color="#0000FF"><strong><br>
        图片:</strong></font>这个要指定好图片所在的位置,这里不可以有其它符号,不要使用;及'这2个符号.<br>
        <strong><font color="#0000FF"><br>
        攻击率:</font></strong>rnd&lt;1/攻击率,当为1时npc攻击为100%,值越大攻击机会就越小.除此外,进行此判断必须是在npc攻击时间内.<br>
        <br>
        <font color="#0000FF"><strong>出现率:</strong></font>死亡时间与当前日期的差如果大于此值,这个npc可以被添到到系统中.越小,npc现出机会越大.单位是秒.<br>
        <br>
        <font color="#0000FF"><strong>物品:</strong></font> npc死后的物品,格式为:出现率1|物品名1,出现率2|物品名2,要说明的是这里的出现率,他的值:<br>
        (死次数/出现率)=取整(死次数/出现率) 如果相同,则出现物吕,比发设置成20,只有NPC每死20次才会有物品出现,设置成1则是死一次都会出现.<br>
        <br>
        <strong><font color="#0000FF">招式:</font></strong>Npc攻击时使用的招式,如果内力不足时,他会使用基本招式进行攻击,比如咬,抓,踢等.这3个基本招式在数据库中各万不要删除.否则无法使用.<br>
        <br>
        <font color="#0000FF"><strong>出场词:</strong></font>Npc现出时的话的内容,这里不要出现&quot;及'号.所以请大家注意.<br>
        <br>
        
      </td>
    </tr>
  </table>
</form>
<br>
</body>
</html>
