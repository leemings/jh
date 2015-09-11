<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
if Application("jhshowurl")<>"http://qqshow-item.tencent.com" then
	showgif=Application("jhshowurl") & "gif/pic/[img].gif"
	showgif1="src="& Application("jhshowurl") & "gif/'+j+'/'+showArray[i]+'.gif"
else
	showgif=Application("jhshowurl") & "/[img]/0/00/00.gif"
	showgif1="src="& Application("jhshowurl") &"/'+showArray[i]+'/'+j+'/00/00.gif"
end if
name=request("name")
id=trim(request("id"))
sex=trim(request("sex"))
if sex<>"男" and sex<>"女" then sex="男"
if (not isnumeric(id)) or instr(name,chr(39))<>0 then 
	Response.Write "<script language=JavaScript>{alert('提示：数据提交有误！');window.close();}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from use where a='"& name &"'", conn, 1,1
if rs.eof and rs.bof then
	if sex="男" then
		Response.Write "var jhshow='||||||14|13|12||11||10|9||||8|||||||';"
	else
		Response.Write "var jhshow='||||||7|6|5||4||3|2||||1|||||||';"
	end if
else
	if id=0 then
		Response.Write "var jhshow='"& rs("f") &"';"
	else
		myhpoto=rs("g")
		showdata=Split(myhpoto, ";")
		if ubound(showdata)>id then
			Response.Write "var jhshow='"& rs("f") &"';"
		else
			Response.Write "var jhshow='"& showdata(id-1) &"';"
		end if
	end if
end if
randomize
myrnd=int(rnd*999)+1000
'Response.Write "document.write ("& chr(34) &"<div id=show"& myrnd &" style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; LEFT: 0px; PADDING-BOTTOM: 0px; WIDTH: 140px; PADDING-TOP: 0px; POSITION: relative; TOP: 0px; HEIGHT: 226px;'></div>"& chr(34) &");"
Response.Write "var showArray = jhshow.split('|');"
Response.Write "var s='';"
Response.Write "for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
Response.Write "s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';"
Response.Write "}}"
Response.Write "s+='<IMG src=../jhshow/images/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &">';"
Response.Write "document.write (s);"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
