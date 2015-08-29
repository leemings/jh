<!--#include file="conn.asp"-->
<!--#include file="z_fy_conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
 response.buffer=true
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_fyaction.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
stats="法院审判庭"
call nav()
call head_var(0,0,"社区法院","z_fy_fayuan.asp")
call getfyconfig()
if not master and not fymaster and not jymaster and request("faction")<>"baoshi" and request("faction")<>"ceshu" then
  Errmsg=Errmsg+"<br>"+"<li>您没有进入本社区法院审判庭的权限，或使用了非法参数。"
  call dvbbs_error()
else
  response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>法院审判庭</b></td></tr><tr><td valign=middle class=tablebody2 height=100><CENTER>"
  dim id,fyrs,fmx,fbgxx,fygxx,fmz,result,bg,my,fgc,fgp,fysql,fjtype
  id=request("id")
  select case checkstr(request("faction"))
    case "alldel"
      call alldel()
    case "allout"
      call allout()
    case "baoshi"
      call baoshi()
       fyrs.close
       set fyrs=nothing
    case "ceshu"
      if not isInteger(request("id")) then response.end
      call ceshu()
    case "peishen"
      if not isInteger(request("id")) then response.end
       call peishen()
    case else
      result=checkstr(request("result"))
      bg=checkstr(request("bg"))
      my=checkstr(request("my"))
      if isnull(request.form("panci")) or checkstr(request.form("panci"))="无" then
        fgc="未予置评"
      else
        fgc=checkstr(request.form("panci"))
      end if
      fgp=checkstr(request.form("play1"))
      id=checkstr(request.querystring("id"))
      if not isInteger(id) then response.end
      select case checkstr(Request.form("action"))
        case ""
           main()
           fyrs.close
           set fyrs=nothing
        case "agree"
           fmz="1"
           agree()
        case "noagree"
           noagree()
        case "pingfan"
           pingfan()
        case "wugao"
           wugao()
        case "wugaohigh"
           wugaohigh()
        case "del"
           del()
        case "fgagree"
           result=fgp
           fmz="6"
           agree()
      end select
  end select
  connfy.close
  set connfy=nothing
  response.write "</CENTER></td></tr></table>" 
end if
call footer() 
'=========================================================
' Sub: main
' Readme:法官的正常审判职能
'=========================================================
Sub main()
fysql="SELECT * FROM fy WHERE ID=" & id
set fyrs=connfy.execute(fysql)
if fyrs.eof or fyrs.bof then
  Errmsg=Errmsg+"<br>"+"<li>没有找到这个诉状！"
  call dvbbs_error()
else
  if fymaster and fyrs(7)<>"N" then
    Errmsg=Errmsg+"<br>"+"<li>对不起，特邀法官没有终审权利！"
    call dvbbs_error()
    exit sub
  end if
  response.write "<div align=center><br><b>案件编号<font color=red>"&id&"</font> </b><br></div>"
  response.write "<table width=97% border=0 cellspacing=1 cellpadding=3 align=center class=tableborder1>"
  response.write "<tr><td height=24 class=tablebody2>原  告 <a href=dispuser.asp?name="&fyrs(4)&" target=_blank><b>"&fyrs(4)&"</b></a>&nbsp;愤怒地写道："&fyrs(1)&"</td></tr>"
  response.write "<tr><td class=tablebody1><br>"&checkstr(fyrs(3))&"<br><br>对于 <a href=dispuser.asp?name="&left(fyrs(2),10)&" target=_blank><b>"&left(fyrs(2),10)&"</b></a> 这种恶劣行径，希望给于 <b><font color=#FF0000>"&fyrs(5)&"</font></b> 的处罚。</td></tr>"
  response.write "<tr><td height=24 class=tablebody2>被  告 <a href=dispuser.asp?name="&fyrs(2)&" target=_blank><b>"&fyrs(2)&"</b></a>&nbsp;辩解道： </td></tr>"
  response.write "<tr><td class=tablebody1><br>"&checkstr(fyrs(9))&"<br></td></tr>"
  response.write "<tr><td height=24 align=left class=tablebody2><form action=z_fy_fyaction.asp?result="&fyrs(5)&"&id="&id&"&bg="&fyrs(2)&"&my="&fyrs(4)&" method = post> 法官 <b>"&membername&"</b> 判决如下：</td></tr>"
  response.write "<tr><td height=24 align=left class=tablebody1>附言：<input name=panci type=text size=80 value="&fyrs(10)&">"
  select case  fyrs(7)
       case "N" 
         response.write "等您判罚"
       case "B"
         response.write "已判原告败诉"
       case "P"
         response.write "原案已平反"
       case "C"
         response.write "原告已撤诉"
       case else
         response.write "已进行 <font color=red><b>"&fyrs(7)&"</b></font> 的处罚"
  end select
  response.write "</td></tr>"
  response.write "<tr><td height=24 align=left class=tablebody2><b>判决选项：</b>"
  response.write "<select name=action size=1><option value='agree'><b>同意原告</b></option><option value='noagree'><b>驳回起诉</b></option>"
  response.write "<option value='pingfan'><b>被告平反</b></option><option value='wugao'><b>原告诬告</b></option>"
  response.write "<option value='wugaohigh'><b>严重诬告</b></option><option value='del'><b>删除本状</b></option><option value='fgagree'><b>【重新量刑】</b></option></select>&nbsp;&nbsp;"
  response.write "<b>重新量刑尺度：</b><select name='play1' size=1>"
  %>
                <option value="终身服刑"><b>╋终身服刑</b></option>
                <option value="入狱10分钟"><b>&nbsp;&nbsp;├入狱10分钟</b></option>
                <option value="入狱半小时"><b>&nbsp;&nbsp;├入狱半小时</b></option>
                <option value="入狱一小时"><b>&nbsp;&nbsp;├入狱一小时</b></option>
                <option value="入狱一天"><b>&nbsp;&nbsp;├入狱一天</b></option>
                <option value="入狱三天"><b>&nbsp;&nbsp;├入狱三天</b></option>
                <option value="入狱一周"><b>&nbsp;&nbsp;├入狱一周</b></option>
                <option value="没收全部财产"><b>╋≡罚收全部财产≡</b></option>               
                <option value="罚款1％"><b>&nbsp;&nbsp;├罚款1%</b></option>
                <option value="罚款10％"><b>&nbsp;&nbsp;├罚款10%</b></option>
                <option value="罚款50％"><b>&nbsp;&nbsp;├罚款50%</b></option>
                <option value="扣除所有经验"><b>╋≡扣除所有经验≡</b></option>               
                <option value="扣经验1％"><b>&nbsp;&nbsp;├扣经验1%</b></option>
                <option value="扣经验10％"><b>&nbsp;&nbsp;├扣经验10%</b></option>               
                <option value="扣经验50％"><b>&nbsp;&nbsp;├扣经验50%</b></option>
                <option value="扣除所有魅力"><b>╋≡扣除所有魅力≡</b></option> 
                <option value="扣魅力1％"><b>&nbsp;&nbsp;├扣魅力1%</b></option>               
                <option value="扣魅力10％"><b>&nbsp;&nbsp;├扣魅力10%</b></option>               
                <option value="扣魅力50％"><b>&nbsp;&nbsp;├扣魅力50%</b></option>
                <option value="清空"><b>╋≡所有分值清零≡</b></option> 
                <option value="威望-10"><b>&nbsp;&nbsp;├威望-10</b></option>
                <option value="威望-50"><b>&nbsp;&nbsp;├威望-50</b></option>
                <option value="威望-100"><b>&nbsp;&nbsp;├威望-100</b></option>
                <option value="午门斩首"><b>╋≡午门斩首≡</b></option>
  <%
  if lh_set(0)="1" then response.write "<option value='离婚'><b>╋≡与被告离婚≡</b></option>"
  response.write "</select>&nbsp;&nbsp;<input type=checkbox name='giveto'  value='1'>罚没积分给胜方(仅在判罚扣减积分时有效）"
  response.write "</td></tr><tr><td class=tablebody1 align=left><b>入狱选项：</b><input type=checkbox name='nobaoshi'  value='1'>禁止保释&nbsp;&nbsp;自定保释金:<input name=bsmoney type=text size=10 value="&bs_set(1)&"> （以上仅在判罚入狱时有效)</td></tr>"
  response.write "<tr><td class=tablebody2 height=24 align=center><input type=submit name=Submit value='执 行'>&nbsp;&nbsp;<input type=button value='返 回' onClick='javascript:history.back()' name=button></td></tr></table>"
end if
End sub
'=========================================================
' Sub: del
' Readme:法官正常审判-删除单个状纸
'=========================================================
Sub del()
  sql="delete from fy where id=" & id & ""
  set fyrs=connfy.execute(sql)
  'fyrs.close
  set fyrs=nothing
  Response.Redirect "z_fy_over.asp"
End sub
'=========================================================
' Sub: alldel
' Readme:状纸列表处理-删除已审结案件
'=========================================================
Sub alldel()
if master then
  sql="delete from fy where result<>'N'"
  connfy.execute sql
end if
  Response.Redirect "z_fy_over.asp"
End sub
'=========================================================
' Sub: allout
' Readme:监狱大赦天下-释放所有犯人
'=========================================================
Sub allout()
if master then
  sql="update [user] set lockuser=0 where lockuser=1"
  conn.execute sql
  fysql="update fy set stats='7',dateandtime=now() where result='入狱三天' or result='入狱一周' or result='入狱一月' or result='终身服刑'"
  connfy.execute fysql
  call bnew("社区监狱实行大赦","释放监狱中关押的所有犯人",membername)
end if
  Response.Redirect "z_fy_fanren.asp"
End sub
'=========================================================
' Sub: baoshi
' Readme:监狱-保释某一犯人
'=========================================================
Sub baoshi()
dim mysign_set
  if bs_set(0)="0" then
    Errmsg=Errmsg+"<br>"+"<li>社区监狱目前不允许保释！"
    call dvbbs_error()
    exit sub
  end if
  sql="select top 1 sign from log where class='监狱' and title='入狱' and tousername='"&request("name")&"' ORDER BY dateandtime DESC"
  set fyrs=connfy.execute(sql)
  if fyrs.eof then
    Errmsg=Errmsg+"<br>"+"<li>没有该犯人的记录，不能保释！"
    call dvbbs_error()
    exit sub
  end if
  mysign_set=split(fyrs(0),"|")
  if not jymaster and not master then
    if mysign_set(0)="1" then
       Errmsg=Errmsg+"<br>"+"<li>该犯人不能保释，只能由监狱长释放！"
       call dvbbs_error()
       exit sub
    end if
    sql="select userwealth from [user] where username='"&membername&"'"
    set rs=conn.execute(sql)
    if clng(rs(0))<=clng(mysign_set(1)) then
      Errmsg=Errmsg+"<br>"+"<li>您的保释金不够！您必须有"+mysign_set(1)+"元才能保释该犯人"
      call dvbbs_error()
      exit sub
    end if
    sql="update [user] set userwealth=userwealth-"&bs_set(1)&" where  username='"&membername&"'"
    conn.execute sql
    content="保释在押犯人 "&request("name")&"，花费了"&mysign_set(1)&"元"
    mysign="0|0"
    call logs("监狱","保释",membername,request("name"))
  else
    mysign="0|0"
    content="鉴于"&request("name")&"在押期间表现良好，被典狱长释放"
    call logs("监狱","出狱",membername,request("name"))
  end if
  sql="update [user] set lockuser=0 where lockuser=1 and username='"&request("name")&"'"
  conn.execute sql
  Response.Redirect "z_fy_fanren.asp"
End sub
'=========================================================
' Sub: ceshu
' Readme:原告撤诉处理
'=========================================================
Sub ceshu()
  sql="select yg,result from fy where id=" & id & ""
  set fyrs=connfy.execute(sql)
  if membername<>fyrs(0) or fyrs(1)<>"N" then
    Errmsg=Errmsg+"<br>"+"<li>您没有撤销此诉状的权限或该案已结案！"
    call dvbbs_error()
    exit sub
  end if
  fysql="update fy set stats='8',result='C',dateandtime=now() where id=" & id & ""
  connfy.execute(fysql)
  fyrs.close
  set fyrs=nothing
  Response.Redirect "z_fy_over.asp"
End sub
'=========================================================
' Sub: peishen()
' Readme:提交陪审团审议
'=========================================================
Sub peishen()
	if pst_set(0)="0" then
		Errmsg=Errmsg+"<br>"+"<li>陪审团审议功能目前被关闭，如想使用请打开并设置相应投票版面！"
		call dvbbs_error()
		exit sub
	else
		boardid=trim(pst_set(1))
		fysql="select * from fy where id="&id
		set fyrs=connfy.execute(fysql)
		response.write "<form action=Savevote.asp?boardID="&boardid&" method=POST onSubmit=submitonce(this) name=frmAnnounce>"
		response.write "<input type=hidden name=upfilerename>"
		response.write "<script src='inc/ubbcode.js'></script>"
		response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"
		response.write "<tr><th width=100% colspan=2 align=center>&nbsp;&nbsp;将法字第 "&id&" 号案件送交陪审团审议"
		if isaudit=1 then response.write "（本版面所有帖子都要经过管理员审核方可发表）"
		response.write "</th></tr>"
		response.write "<tr><td width=20% class=tablebody2><b>用户名：</b></td>"
		response.write "<td width=80% class=tablebody2><input name=username value="&membername&" class=FormClass disabled>&nbsp;&nbsp;<b>密码：</b><input name=passwd type=password value="&htmlencode(memberword)&" class=FormClass disabled></td></tr>"
		response.write "<tr><td width=20% class=tablebody1><b>标题：</b></td>"
		response.write "<td width=80% class=tablebody2><input name=subject size=50 maxlength=80 class=FormClass value='[法字第"&id&"号案件]请各位参与审议'>&nbsp;<select name=votetimeout size=1>"
		response.write "<option value=0>审议时间</option><option value=0>永久审议</option><option value=1>一天</option><option value=3>三天</option><option value=7>一周</option><option value=15>半月</option></select>&nbsp;<b>*</b>标题不得超过25个汉字或50个英文字符"
		response.write "</td></tr>"
		response.write "<tr><td width=20% class=tablebody2><b>判罚选项： </b> <br><br><li>每行一个判罚项目</li><li>法官可根据具体情况修改</li><br><br>"
		response.write "<input type=radio name=votetype value=0 checked >单选投票<br></font></td>"
		response.write "<td width=80% class=tablebody2><textarea name=vote cols=95 rows=8>同意原告（按原告意见处罚被告）"&chr(13)&"同意原告（并对被告加重处罚）"&chr(13)&"同意原告（可对被告减轻处罚）"&chr(13)&"原告纯属造谣，应受处罚"&chr(13)&"原告纯属诬告，应受重罚</textarea>"
		response.write "</td></tr>"
		response.write "<tr><td width=20% valign=top class=tablebody1><b>当前心情：</b><br><li>将放在帖子的前面</td><td width=80% class=tablebody1>"
		for i=0 to Forum_PostFaceNum-1
			response.write "<input type=radio value='"&Forum_PostFace(i)&"' name='Expression'"
			if i=0 then response.write "checked"
			response.write "><img src="&forum_info(8)+Forum_PostFace(i)&" >&nbsp;&nbsp;"
			if i>0 and ((i+1) mod 9=0) then response.write "<br>"
		next
		response.write "</td></tr>"
		response.write "<tr><td width=20% valign=top class=tablebody2><b>案件详情：</b><br></td><td width=80% class=tablebody2><br>"
		response.write "<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL title=可以使用Ctrl+Enter直接提交帖子 class=FormClass onkeydown=ctlent()>[b]投诉日期：[/b]"&fyrs(8)&chr(13)
		response.write "[b]被告：[/b]"&fyrs(2)&chr(13)&"[b]原告：[/b]"&fyrs(4)&chr(13)&"[b]原告要求对被告处以"&fyrs(5)&"的惩罚[/b]"&chr(13)&chr(13)&"[b]投诉标题：[/b]"&fyrs(1)&chr(13)&"[b]投诉详情：[/b]"&fyrs(3)&chr(13)&"[b]被告申诉：[/b]"&fyrs(9)&chr(13)&chr(13)&"[color=#FF0000][url=z_fy_fyaction.asp?id="&id&"]法官请点这里进行最终审判[/url][/color]</textarea>"
		response.write "</td></tr>"
		response.write "<tr><td valign=middle colspan=2 align=center class=tablebody2><input type=Submit value='发 表' name=Submit> &nbsp; <input type=button value='预 览' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value='清 除'></td></form></tr></table></form>"
		response.write "<form name=preview action=preview.asp?boardid="&boardid&" method=post target=preview_page><input type=hidden name=title value=><input type=hidden name=body value=></form>"
	end if
	fyrs.close
	set fyrs=nothing
End sub
'=========================================================
' Sub: fmcl()
' Readme:普通处罚-同意原告积分处理
'=========================================================
sub fmcl()
sql="select "&fmx&"*"&fygxx&" from [user] where username='"&bg&"'"
set rs=conn.execute(sql)
if request.form("giveto")="1" then
 sql="update [user] set "&fmx&"="&fmx&"+round("&rs(0)&") where username='"&my&"'"  
 conn.execute sql
end if
fysql="update fy set lastvalue=round("&rs(0)&") where id=" & id & ""
connfy.execute fysql
sql="update [user] set "&fmx&"=round("&fmx&"*"&fbgxx&") where username='"&bg&"'"  
conn.Execute(sql)
end sub
'=========================================================
' Sub: fjfj()
' Readme:附加罚金处理
'=========================================================
sub fjfj()
select case fjtype
case "1"
  if fmz="1" and fjty_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjty_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fmz="1" and fjty_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjty_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "2"
  if fjbh_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjbh_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjbh_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjbh_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "3"
  if fjpf_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjpf_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjpf_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjpf_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "4"
  if fjzy_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjzy_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjzy_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjzy_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
case "5"
  if fjwg_set(0)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjwg_set(0)&" where username='"&my&"'"  
    conn.execute(sql)
  end if
  if fjwg_set(1)<>"0" then
    sql="update [user] set userwealth=userwealth"&fjwg_set(1)&" where username='"&bg&"'"  
    conn.execute(sql)
  end if
end select
end sub
'=========================================================
' Sub: agree()
' Readme:同意原告/法官自定义判罚
'=========================================================
sub agree()
fjtype="1"
select case result
case "罚款1％" 
  fmx="userwealth"
  fbgxx="0.99"
  fygxx="0.01"
  call fmcl()
case "罚款10％" 
  fmx="userwealth"
  fbgxx="0.9"
  fygxx="0.1"
  call fmcl()
case "罚款50％" 
  fmx="userwealth"
  fbgxx="0.5"
  fygxx="0.5"
  call fmcl()
case "没收全部财产"
  fmx="userwealth"
  fbgxx="0"
  fygxx="1"
  call fmcl()
case "扣经验1％" 
  fmx="userep"
  fbgxx="0.99"
  fygxx="0.01"
  call fmcl()
case "扣经验10％"
  fmx="userep"
  fbgxx="0.9"
  fygxx="0.1"
  call fmcl()
case "扣经验50％"
  fmx="userep"
  fbgxx="0.5"
  fygxx="0.5"
  call fmcl()
case "扣除所有经验" 
  fmx="userep"
  fbgxx="0"
  fygxx="1"
  call fmcl()
case "扣魅力1％"
  fmx="usercp"
  fbgxx="0.99"
  fygxx="0.01"
  call fmcl()
case "扣魅力10％"
  fmx="usercp"
  fbgxx="0.9"
  fygxx="0.1"
  call fmcl()
case "扣魅力50％"
  fmx="usercp"
  fbgxx="0.5"
  fygxx="0.5"
  call fmcl()
case "扣除所有魅力" 
  fmx="usercp"
  fbgxx="0"
  fygxx="1"
  call fmcl()
case "威望-10"
  sql="update [user] set userpower=userpower-10 where username='" & bg & "'"
  conn.execute sql
  if request.form("giveto")="1" then
    sql="update [user] set userpower=userpower+10 where username='"&my&"'"  
    conn.execute sql
   end if
case "威望-50"
  sql="update [user] set userpower=userpower-50 where username='" & bg & "'"
  conn.execute sql
  if request.form("giveto")="1" then
    sql="update [user] set userpower=userpower+50 where username='"&my&"'"  
    conn.execute sql
   end if
case "威望-100" 
  sql="update [user] set userpower=userpower-100 where username='" & bg & "'"
  conn.execute sql
   if request.form("giveto")="1" then
    sql="update [user] set userpower=userpower+100 where username='"&my&"'"  
    conn.execute sql
   end if
case "清空" 
  sql="update [user] set userwealth=0,usercp=0,userep=0,userpower=0 where username='" & bg & "'"
  conn.execute sql
case "入狱10分钟" 
   content="由于违反社区规定被判入狱10分钟"
   call jjy()
case "入狱半小时" 
   content="由于违反社区规定被判入狱半小时"
   call jjy()
case "入狱一小时" 
   content="由于违反社区规定被判入狱一小时"
   call jjy()
case "入狱一天" 
   content="由于违反社区规定被判入狱一天"
   call jjy()
case "入狱三天"  
  content="由于违反社区规定被判入狱三天"
  call jjy()
case "入狱一周" 
  content="由于违反社区规定被判入狱一周"
  call jjy()
case "终身服刑"
  content="由于违反社区规定被判终身服刑"
  call jjy()
case "午门斩首" 
  sql="delete from [user] where username='" & bg & "'"
  conn.execute sql
case "离婚"
  if lh_set(0)="0" then
    Errmsg=Errmsg+"<br>"+"<li>离婚判罚功能目前是关闭的！"
    call dvbbs_error()
    exit sub
  end if
  if lh_set(0)="1" then
    '以下代码由lwand提供
    sql="update [user] set wife=' ' where username='" & bg & "'"' or username='" & my "' "
    conn.execute(sql)
    connj.execute("delete from jie where username='" & my & "'")
    connj.execute("delete from jie where username='" & bg & "'")
    connj.execute("delete from qiuhun where username='" & my & "'")
    connj.execute("delete from qiuhun where username='" & bg & "'")
    '以上代码由lwand提供
    if lh_set(1)="1" then
      sql="update [user] set usercp=usercp-"&lh_set(2)&" where username='"&bg&"'"  
      conn.execute(sql)
    end if
  end if
end select
if founderr then 
 call dvbbs_error()
 exit sub
end if
call fjfj()
fysql="update fy set stats='"&fmz&"',fgtext='"&fgc&"',result='"&result&"',dateandtime=now() where id=" & id & ""
connfy.execute fysql
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: noagree()
' Readme:驳回原告
'=========================================================
sub noagree()
fjtype="2"
fysql="update fy set stats='0',fgtext='"&fgc&"',result='B',dateandtime=now() where id=" & id & ""
response.write fysql
connfy.execute fysql
call fjfj()
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: pingfan()
' Readme:平反
'=========================================================
sub pingfan()
fjtype="3"
call getlast()
if fyrs(0)="罚款1％" or fyrs(0)="罚款10％" or fyrs(0)="罚款50％" or fyrs(0)="没收全部财产" then
  sql="update [user] set userwealth=userwealth+"&fyrs(1)&" where username='"&bg&"'"  
  conn.Execute(sql)
end if
if fyrs(0)="扣经验1％" or fyrs(0)="扣经验10％" or fyrs(0)="扣经验50％" or fyrs(0)="扣除所有经验" then
    sql="update [user] set userep=userep+"&fyrs(1)&" where username='" & bg & "'"
    conn.execute sql
end if
if fyrs(0)="扣魅力1％" or fyrs(0)="扣魅力10％" or fyrs(0)="扣魅力50％" or fyrs(0)="扣除所有魅力" then
    sql="update [user] set usercp=usercp+"&fyrs(1)&" where username='" & bg & "'"
    conn.execute sql
end if
if fyrs(0)="威望-10" then
sql="update [user] set userpower=userpower+10 where username='" & bg & "'"
conn.execute sql
end if
if fyrs(0)="威望-50" then
sql="update [user] set userpower=userpower+50 where username='" & bg & "'"
conn.execute sql
end if
if fyrs(0)="威望-100" then
sql="update [user] set userpower=userpower+100 where username='" & bg & "'"
conn.execute sql
end if
if fyrs(0)="清空" or fyrs(0)="N" or fyrs(0)="P" or fyrs(0)="午门斩首" or fyrs(0)="C" or fyrs(0)="离婚" then 
   Errmsg=Errmsg+"<br>"+"<li>该案件尚未审判或已撤诉或此判决无法平反！"
   call dvbbs_error()
   exit sub
end if
if fyrs(0)="入狱10分钟" or fyrs(0)="入狱半小时" or fyrs(0)="入狱一小时" or fyrs(0)="入狱一天" or fyrs(0)="入狱三天" or fyrs(0)="入狱一周" or fyrs(0)="终身服刑" then
   sql="update [user] set lockuser=0 where username='" & bg & "'"
   conn.execute sql
   mysign="0|0"
   content="案件被平反，得以出狱"
   call logs("监狱","出狱",membername,bg)
end if
if fyrs(0)="B" then
  sql="update [user] set userwealth=userwealth+"&fyrs(1)&" where username='"&my&"'"  
  conn.Execute(sql)
end if
fysql="update fy set stats='2',fgtext='"&fgc&"',result='P',dateandtime=now() where id=" & id & ""
connfy.execute fysql
call fjfj()
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: wugao()
' Readme:诬告
'=========================================================
sub wugao()
fjtype="4"
call fjfj()
fysql="update fy set stats='4',fgtext='"&fgc&"',result='B',dateandtime=now(),lastvalue="&fjzy_set(0)&" where id=" & id & ""
connfy.execute fysql
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: wugaohigh()
' Readme:严重诬告
'=========================================================
sub wugaohigh()
fjtype="5"
call fjfj()
fysql="update fy set stats='5',fgtext='"&fgc&"',result='B',dateandtime=now(),lastvalue="&fjzy_set(0)&" where id=" & id & ""
connfy.execute fysql
Response.Redirect "z_fy_over.asp"
end sub
'=========================================================
' Sub: getlast()
' Readme:读取上次被罚积分，用于平反
'=========================================================
sub getlast()
fysql="select result,lastvalue from fy where id=" & id & ""
set fyrs=connfy.execute(fysql)
end sub
'=========================================================
' Sub: jjy()
' Readme:入狱处理
'=========================================================
sub jjy()
dim bbsmoney
if not isInteger(request.form("bsmoney")) or request.form("bsmoeny")<0 then
   Errmsg=Errmsg+"<br>"+"<li>您设定的保释金不合法！"
   founderr=true
   exit sub
else
if clng(request.form("bsmoney"))<=clng(bs_set(1)) then
  bbsmoney=bs_set(1)
else
  bbsmoney=request.form("bsmoney")
end if
  sql="update [user] set lockuser=1 where username='" & bg & "'"
  conn.execute sql
  if request.form("nobaoshi")="1" then
    mysign="1" & "|" & bbsmoney
  else
    mysign="0" & "|" & bbsmoney
  end if
  call logs("监狱","入狱",membername,bg)
end if
end sub
%>
