<%dim LocalPic,PicPath,CartLength,hasVipPlus,DefaultVisualBoy,DefaultVisualGirl,isAllow,DefaultFont,PhotoPath,PhotoListCount,DefaultStatus

hasVipPlus=1	'如果你没有安装VIP插件，请设置为0
CartLength=10		'购物袋的大小
LocalPic=0		'0-使用腾讯图片 1-使用本地图片
isAllow=1			'是否允许不雅形象出现，0-否，1-是
DefaultFont="宋体"		'默认字体
PhotoPath="UserVisual"	'保存形象照片的目录
PhotoListCount=5
DefaultStatus=conn.execute("Select top 1 PhotoStatus from Visual_config")(0) 		' 个人写真的状态，0-秘密，1-公开
DefaultVisualBoy="||||||14.7|13.8|12.9||11.11||10.13|9.14||||8.18|||||||"
DefaultVisualGirl="||||||7.7|6.8|5.9||4.11||3.13|2.14||||1.18|||||||"
if LocalPic=1 then
	PicPath="Images/Img_Visual/Show/"
else
	PicPath="http://qqshow-item.tencent.com/"
end if

if right(lcase(scriptname),11)<>"dispbbs.asp" then
Dim v_mysex,v_myvip,v_myface,v_myvisual,v_mywidth,v_myheight
v_mysex=1
v_myvip=0
v_myface="userface/image1.gif"
v_myvisual=""
v_mywidth=forum_setting(38)
v_myheight=forum_setting(39)
if founduser then
	set rs=server.createobject("adodb.recordset")
	if hasVipPlus=1 then
		sql="select sex,vip,face,visual,width,height from [user] where userid="&userid
	else
		sql="select sex,face,visual,width,height from [user] where userid="&userid
	end if
	rs.open sql,conn,1,1
	if not (rs.eof or rs.bof) then
		v_mysex=rs("sex")
		if hasVipPlus=1 then
			v_myvip=rs("vip")
		end if
		if isnull(v_myvip) then
			v_myvip=0
		end if
		v_myface=htmlencode(FilterJS(rs("face")))
		if isnull(rs("width")) or rs("width")=0 then
			v_mywidth=forum_setting(38)
		else
			v_mywidth=rs("width")
		end if
		if isnull(rs("height")) or rs("height")=0 then
			v_myheight=forum_setting(39)
		else
			v_myheight=rs("height")
		end if
		if session("period_"&userid)="" or not isdate(session("period_"&userid)) then
			v_myvisual=GetVisualStr(rs("visual"),v_mysex)
			conn.execute("update visual_myitems set type=0 where type<>0 and type<>5 and datediff('d',fixdate,now())>period and period>0 and username='"&membername&"'")
			call DeletePhoto()
			session("period_"&userid)=now()
		elseif datediff("d",session("period_"&userid),now())>0 then
			v_myvisual=GetVisualStr(rs("visual"),v_mysex)
			conn.execute("update visual_myitems set type=0 where type<>0 and type<>5 and datediff('d',fixdate,now())>period and period>0 and username='"&membername&"'")
			call DeletePhoto()
			session("period_"&userid)=now()
		else
			if isnull(rs("visual")) or rs("visual")="" then
				if v_mySex=1 then
					v_myvisual=DefaultVisualBoy
				else
					v_myvisual=DefaultVisualGirl
				end if
			else
				v_myvisual=rs("visual")
			end if
		end if
		if v_myvisual<>rs("visual") then conn.execute("update [user] set visual='"&v_myvisual&"' where userid="&userid)
	end if
	rs.close
	set rs=nothing
end if
end if

sub DeletePhoto()
	dim rsPhoto
	set rsPhoto=server.createobject("adodb.recordset")
	sql="select id from visual_photo where enddate<now() and fromuser='"&membername&"' and not finished"
	rsphoto.open sql,conn,1,1
	do while not rsphoto.eof
		conn.execute("delete from visual_photouser where photo_id="&rsphoto(0)) 
		rsphoto.movenext
	loop
	set rsphoto=nothing
	conn.execute("delete from visual_photo where enddate<now() and fromuser='"&membername&"' and not finished") 
end sub

function GetVisualStr(CurVisualStr,CurVisualSex)
	dim VisualStr,VisualSex
	dim VisualSplit,rsPeriod
	dim i,TempSplit
	VisualStr=CurVisualStr
	VisualSex=CurVisualSex
	if isnull(VisualStr) or VisualStr="" then
		if VisualSex=1 then
			VisualStr=DefaultVisualBoy
		else
			VisualStr=DefaultVisualGirl
		end if
	else
		if ubound(split(VisualStr,"|"))<>24 then
			if VisualSex=1 then
				VisualStr=DefaultVisualBoy
			else
				VisualStr=DefaultVisualGirl
			end if
		else
			VisualSplit=split(VisualStr,"|")
			VisualStr=""
			set rsPeriod=server.createobject("ADODB.recordset")
			for i=0 to 24
				if VisualSplit(i)<>"" then
					TempSplit=split(VisualSplit(i),".")
					if cint(TempSplit(0))>20 then
						sql="select fixdate,Period from visual_myitems where itemid="&TempSplit(0)&" and username='"&membername&"' and ((datediff('d',fixdate,now())<=period and period>0) or period=0)"
						rsPeriod.open sql,conn,1,1
						if rsPeriod.eof or rsPeriod.bof then
							select case i
							case 6
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(6)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(6)
								end if
							case 7
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(7)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(7)
								end if
							case 8
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(8)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(8)
								end if
							case 10
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(10)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(10)
								end if
							case 12
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(12)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(12)
								end if
							case 13
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(13)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(13)
								end if
							case 17
								if VisualSex=1 then
									VisualSplit(i)=split(DefaultVisualBoy,"|")(17)
								else
									VisualSplit(i)=split(DefaultVisualGirl,"|")(17)
								end if
							case else
								VisualSplit(i)=""
							end select
						end if
						rsPeriod.close
					end if
				end if
				if i>0 then VisualStr=VisualStr&"|"
				VisualStr=VisualStr&VisualSplit(i)
			next
			set rsPeriod=nothing
		end if
	end if
	GetVisualStr=VisualStr
end function

function GetLayerStr(CurVisualSplit)
	dim i,LayerStr
	LayerStr=""
	for i=0 to ubound(CurVisualSplit)
		if not isnull(CurVisualSplit(i)) and trim(CurVisualSplit(i))<>"" then
			if LayerStr<>"" then LayerStr=LayerStr&"_"
			LayerStr=LayerStr&(i+1)
		end if
	next
	GetLayerStr=LayerStr
end function

function isCanShowVisual(CurVisualSplit)
	dim TempSplit
	if ubound(CurVisualSplit)<>24 then
		isCanShowVisual=False
		exit function
	end if
	if CurVisualSplit(6)="" or CurVisualSplit(7)="" or CurVisualSplit(8)="" or CurVisualSplit(10)="" or CurVisualSplit(12)="" or CurVisualSplit(13)="" or CurVisualSplit(17)="" then
		isCanShowVisual=False
		exit function
	end if
	if isAllow=0 then
		TempSplit=split(CurVisualSplit(7),".")
		if TempSplit(0)="6" or TempSplit(0)="13" then
			isCanShowVisual=False
			exit function
		end if
		TempSplit=split(CurVisualSplit(8),".")
		if TempSplit(0)="5" or TempSplit(0)="12" then
			isCanShowVisual=False
			exit function
		end if
	end if
	isCanShowVisual=True
end function

sub showvisual(uservisual,visualwidth,usersex,userface,userfacewidth,userfaceheight,bgcolor,isSmall)
	if uservisual<>"" then
		if isCanShowVisual(split(uservisual,"|")) then
			if isSmall or visualwidth<=0 then
				response.write "<CENTER><IFrame src=z_visual_show.asp?width=0&picpath="&picpath&"&localpic="&localpic&"&visual="&uservisual&" width='80' height='120' marginwidth='0' marginheight='0' hspace='0' vspace='0' frameborder='0' scrolling='NO' allowtransparency='true'></iframe></CENTER>"
			else
				response.write "<CENTER><IFrame src=z_visual_show.asp?width="&visualwidth&"&picpath="&picpath&"&localpic="&localpic&"&visual="&uservisual&" width='"&visualwidth&"' height='"&(visualwidth*226\140)&"' marginwidth='0' marginheight='0' hspace='0' vspace='0' frameborder='0' scrolling='NO' allowtransparency='true'></iframe></CENTER>"
			end if
		else
			response.write "&nbsp;&nbsp;<img src="""&userface&""" border=0 width="""&userfacewidth&""" height="""&userfaceheight&""">"
		end if
	else
		response.write "&nbsp;&nbsp;<img src="""&userface&""" border=0 width="""&userfacewidth&""" height="""&userfaceheight&""">"
	end if
end sub

Function GetBody(UserVisual,isPhotoMe)
	dim VisualSplit,TempSplit,i
	dim Result
	if UserVisual="" then
		GetBody=""
		Exit Function
	end if
	VisualSplit=Split(UserVisual,"|")
	if ubound(VisualSplit)<>24 then
		GetBody=""
		Exit Function
	end if
	Result=""
	for i=0 to ubound(VisualSplit)
		if ((i>=4 and i<=20) or isPhotoMe) and i<>1 then
			if VisualSplit(i)<>"" then
				TempSplit=Split(VisualSplit(i),".")
				if Result<>"" then Result=Result&"|"
				Result=Result&(i+1)&"/"&right("0000"&TempSplit(0),4)&".Gif"
			end if
		end if
	next
	GetBody=Result
End Function

Function GetBodyDiv(SeqNo,UserVisual,ZIndex,OuterLeft,OuterTop,OuterWidth,OuterHeight,InnerLeft,InnerTop,InnerWidth,InnerHeight,Direction,isPhotoMe)
	dim VisualSplit,TempSplit,i
	dim Result
	if UserVisual="" then
		GetBodyDiv=""
		Exit Function
	end if
	VisualSplit=Split(UserVisual,"|")
	if ubound(VisualSplit)<>24 then
		GetBodyDiv=""
		Exit Function
	end if
	Result="<DIV id=""BodyBorder_"&SeqNo&""" style=""overflow:hidden;z-index:"&ZIndex&";left:"&OuterLeft&";width:"&OuterWidth&";position:absolute;top:"&OuterTop&";height:"&OuterHeight&";filter:progid:DXImageTransform.Microsoft.BasicImage(rotation=0,mirror="&(Direction\4)&")""><DIV id=""BodyCont_"&SeqNo&""" style=""z-index:"&(ZIndex+1)&";left:"&InnerLeft&";width:"&InnerWidth&";position:absolute;top:"&InnerTop&";height:"&InnerHeight&";"">"
	for i=0 to ubound(VisualSplit)
		if ((i>=4 and i<=20) or isPhotoMe) and i<>1 then
			if VisualSplit(i)<>"" then
				TempSplit=Split(VisualSplit(i),".")
      	if localpic=1 then
					Result=Result&"<img src="""&PicPath&(i+1)&"/"&right("0000"&TempSplit(0),4)&".Gif"" style=""padding:0;position:absolute;top:0;left:0;width:"&InnerWidth&";height:"&InnerHeight&";z-index:"&(ZIndex+i-2)&";"" name=photo_"&SeqNo&">"
      	else
					Result=Result&"<img src="""&PicPath&right("00000000"&TempSplit(0),8)&"/"&right("00"&(i+1),2)&"/00/"" style=""padding:0;position:absolute;top:0;left:0;width:"&InnerWidth&";height:"&InnerHeight&";z-index:"&(ZIndex+i-2)&";"" name=photo_"&SeqNo&">"
      	end if
			end if
		end if
	next
	Result=Result&"<img src=""images/img_visual/blank.gif"" style=""padding:0;position:absolute;top:0;left:0;width:"&InnerWidth&";height:"&InnerHeight&";z-index:"&(ZIndex+20)&";"" name=photo_"&SeqNo&">"
	Result=Result&"</DIV></DIV>"
	GetBodyDiv=Result
End Function

Function GetAccoutermentDiv(SeqNo,UserVisual,ZIndex,OuterLeft,OuterTop,OuterWidth,OuterHeight,InnerLeft,InnerTop,InnerWidth,InnerHeight,Direction)
	dim Result
	Result="<DIV id=""AccoutermentBorder_"&SeqNo&""" style=""overflow:hidden;z-index:"&ZIndex&";left:"&OuterLeft&";width:"&OuterWidth&";position:absolute;top:"&OuterTop&";height:"&OuterHeight&";filter:progid:DXImageTransform.Microsoft.BasicImage(rotation=0,mirror="&(Direction\4)&")""><IMG src="""
	if UserVisual<>"" then
		Result=Result&PicPath&Uservisual
	end if
	Result=Result&""" style=""padding:0;z-index:"&(ZIndex+1)&";left:"&InnerLeft&";width:"&InnerWidth&";position:absolute;top:"&InnerTop&";height:"&InnerHeight&";"
	if UserVisual<>"" then
		Result=Result&"display:inline"
	else
		Result=Result&"display:none"
	end if
	Result=Result&"""></DIV>"
	GetAccoutermentDiv=Result
End Function

Function GetNameDiv(SeqNo,UserName,ZIndex,NameLeft,NameTop,NameFont,NameSize,NameBold,NameItalic,NameColor,NameDirection)
	if UserName="" then
		GetNameDiv=""
		Exit Function
	end if
	GetNameDiv="<DIV id=""NameBorder_"&SeqNo&""" style=""z-index:"&ZIndex&";left:"&NameLeft&";"
	GetNameDiv=GetNameDiv&"position:absolute;top:"&NameTop&";width:140;"
	GetNameDiv=GetNameDiv&"padding-top:3px;padding-left:3px;text-align:left;word-break:keep-all;line-height:"&NameSize&";"
	if NameBold then
		GetNameDiv=GetNameDiv&"font-weight:bolder;"
	else
		GetNameDiv=GetNameDiv&"font-weight:normal;"
	end if
	if NameItalic then
		GetNameDiv=GetNameDiv&"font-style:italic;"
	else
		GetNameDiv=GetNameDiv&"font-style:normal;"
	end if
	GetNameDiv=GetNameDiv&"font-family:"&NameFont&";"
	if NameColor="" then
		GetNameDiv=GetNameDiv&"display:none;color:#FFFFFF;"
	else
		GetNameDiv=GetNameDiv&"color:"&NameColor&";"
	end if
	GetNameDiv=GetNameDiv&"font-size:"&NameSize&";"
	GetNameDiv=GetNameDiv&"filter:progid:DXImageTransform.Microsoft.BasicImage(rotation="&NameDirection&",mirror=0) progid:DXImageTransform.Microsoft.dropShadow(Color=000000,offX=1,offY=1,positive=true)"
	GetNameDiv=GetNameDiv&""">"&UserName&"</div>"
End Function

Function GetPhotoNameDiv(PhotoName,ZIndex,PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection)
	if PhotoName="" then
		GetPhotoNameDiv=""
		Exit Function
	end if
	GetPhotoNameDiv="<DIV id=""PhotoNameBorder"" style=""z-index:"&ZIndex&";left:"&PhotoNameLeft&";"
	GetPhotoNameDiv=GetPhotoNameDiv&"position:absolute;top:"&PhotoNameTop&";width:140;"
	GetPhotoNameDiv=GetPhotoNameDiv&"padding-top:3px;padding-left:3px;text-align:left;word-break:keep-all;line-height:"&PhotoNameSize&";"
	if PhotoNameBold then
		GetPhotoNameDiv=GetPhotoNameDiv&"font-weight:bolder;"
	else
		GetPhotoNameDiv=GetPhotoNameDiv&"font-weight:normal;"
	end if
	if PhotoNameItalic then
		GetPhotoNameDiv=GetPhotoNameDiv&"font-style:italic;"
	else
		GetPhotoNameDiv=GetPhotoNameDiv&"font-style:normal;"
	end if
	GetPhotoNameDiv=GetPhotoNameDiv&"font-family:"&PhotoNameFont&";"
	if PhotoNameColor="" then
		GetPhotoNameDiv=GetPhotoNameDiv&"display:none;color:#FFFFFF;"
	else
		GetPhotoNameDiv=GetPhotoNameDiv&"color:"&PhotoNameColor&";"
	end if
	GetPhotoNameDiv=GetPhotoNameDiv&"font-size:"&PhotoNameSize&";"
	GetPhotoNameDiv=GetPhotoNameDiv&"filter:progid:DXImageTransform.Microsoft.BasicImage(rotation="&PhotoNameDirection&",mirror=0) progid:DXImageTransform.Microsoft.dropShadow(Color=000000,offX=1,offY=1,positive=true)"
	GetPhotoNameDiv=GetPhotoNameDiv&""">"&PhotoName&"</div>"
End Function

Function GetDateDiv(CurDate,ZIndex,DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection)
	GetDateDiv="<DIV id=""DateBorder"" style=""z-index:"&ZIndex&";left:"&DateLeft&";"
	GetDateDiv=GetDateDiv&"position:absolute;top:"&DateTop&";width:140;"
	GetDateDiv=GetDateDiv&"padding-top:3px;padding-left:3px;text-align:left;word-break:keep-all;line-height:"&DateSize&";"
	if DateBold then
		GetDateDiv=GetDateDiv&"font-weight:bolder;"
	else
		GetDateDiv=GetDateDiv&"font-weight:normal;"
	end if
	if DateItalic then
		GetDateDiv=GetDateDiv&"font-style:italic;"
	else
		GetDateDiv=GetDateDiv&"font-style:normal;"
	end if
	GetDateDiv=GetDateDiv&"font-family:"&DateFont&";"
	if DateColor="" then
		GetDateDiv=GetDateDiv&"display:none;color:#FFFFFF;"
	else
		GetDateDiv=GetDateDiv&"color:"&DateColor&";"
	end if
	GetDateDiv=GetDateDiv&"font-size:"&DateSize&";"
	GetDateDiv=GetDateDiv&"filter:progid:DXImageTransform.Microsoft.BasicImage(rotation="&DateDirection&",mirror=0) progid:DXImageTransform.Microsoft.dropShadow(Color=000000,offX=1,offY=1,positive=true)"
	GetDateDiv=GetDateDiv&""">"&formatdatetime(CurDate,2)&"</div>"
End Function

Sub SaveVisual(SaveUserID,SaveUserName,SaveUserClass,SaveUserVisual)
	Dim BodyStr,BodySplit
	Dim pic
	if isNull(SaveUserVisual) or not SaveUserVisual<>"" then
		exit sub
	end if
	if isCanShowVisual(split(SaveUserVisual,"|")) then
		BodyStr=GetBody(SaveUserVisual,false)
		BodySplit=Split(BodyStr,"|")
		on error resume next
		Set pic = CreateObject("ASPPainter.Pictures.1")
		if err.Number=0 then
			on error goto 0
			pic.SetFormat 3
			pic.LoadFile Server.MapPath(PicPath&BodySplit(0))
			for i=1 to ubound(bodysplit)
				pic.SetImageIndex 1
				pic.LoadFile Server.MapPath(PicPath&BodySplit(i))
				pic.Merge 0,1,0,0,0,0,pic.Width,pic.Height,100
				pic.Destroy
				pic.SetImageIndex 0
				pic.SetColor 255,255,255,255
				pic.setColorAsTransparent
			next
			pic.SetFontName "宋体"
			pic.SetFontSize 9
			pic.SetFontOrientation 0
			pic.setTextAlign 1
			if trim(SaveUserName)<>"" then
				pic.SetFontRectangle 1,1,141,16
				pic.SetColor 0,0,0,255
				pic.TextOut 0,0,trim(SaveUserName)
				pic.ClearFontRectangle
				pic.SetFontRectangle 0,0,140,15
				pic.SetColor 255,255,255,255
				pic.TextOut 0,0,trim(SaveUserName)
				pic.ClearFontRectangle
			end if
			if trim(SaveUserClass)<>"" then
				pic.SetFontRectangle 1,16,141,31
				pic.SetColor 0,0,0,255
				pic.TextOut 0,0,trim(SaveUserClass)
				pic.ClearFontRectangle
				pic.SetFontRectangle 0,15,140,30
				pic.SetColor 255,255,255,255
				pic.TextOut 0,0,trim(SaveUserClass)
				pic.ClearFontRectangle
			end if
			if trim(SaveUserName)="" then
				pic.SaveToFile Server.MapPath(PhotoPath&"\Photo_"&SaveUserID&".gif")
			else
				pic.SaveToFile Server.MapPath(PhotoPath&"\"&SaveUserID&".gif")
			end if
			pic.DestroyAll
			Set pic = Nothing
		else
			err.clear
			on error goto 0
		end if
	end if
End Sub

Sub ShowMyselfVisual(v_VisualWidth,v_BodyClass)
	call showvisual(v_myvisual,v_VisualWidth,v_mysex,v_myface,v_mywidth,v_myheight,v_BodyClass,false)
end sub

Sub ShowUserVisual(v_UserName,v_VisualWidth,v_BodyClass,v_isSmallVisual)
	dim v_usersex,v_userface,v_uservisual,v_userwidth,v_userheight,v_rsu
	v_usersex=1
	v_userface="userface/image1.gif"
	v_uservisual=""
	v_userwidth=forum_setting(38)
	v_userheight=forum_setting(39)
	set v_rsu=server.createobject("adodb.recordset")
	sql="select sex,face,visual,width,height from [user] where username='"&v_UserName&"'"
	v_rsu.open sql,conn,1,1
	if not (v_rsu.eof or v_rsu.bof) then
		v_userface=htmlencode(FilterJS(v_rsu("face")))
		if isnull(v_rsu("width")) or v_rsu("width")=0 then
			v_userwidth=forum_setting(38)
		else
			v_userwidth=v_rsu("width")
		end if
		if isnull(v_rsu("height")) or v_rsu("height")=0 then
			v_userheight=forum_setting(39)
		else
			v_userheight=v_rsu("height")
		end if
		v_uservisual=v_rsu("visual")
	end if
	v_rsu.close
	set v_rsu=nothing
	call showvisual(v_UserVisual,v_VisualWidth,v_UserSex,v_UserFace,v_UserWidth,v_UserHeight,v_BodyClass,v_isSmallVisual)
End Sub

sub DispPageNum(CurPage, PageCount, URLPrefix, URLPostfix)
	dim p,ii
	if PageCount=0 then
		response.write "无 "
	else
		p=(CurPage-1) \ 10
		if CurPage=1 then
			response.write "<SPAN class=pagenumstatic><font face=webdings color="&Forum_body(8)&">9</font></SPAN>  "
		else
			response.write "<SPAN class=pagenum><a href="&URLPrefix&"1"&URLPostfix&" title=首页><font face=webdings>9</font></a></SPAN> "
		end if
		if p>0 then response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(p*10)&URLPostfix&" title=上十页><font face=webdings>7</font></a></SPAN> "
		response.write "<b>"
		for ii=p*10+1 to P*10+10
			if ii=CurPage then
		  	response.write "<SPAN class=pagenumstatic><B><font color="&Forum_body(8)&">"+Cstr(ii)+"</font></B></SPAN> "
			else
				response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(ii)&URLPostfix&">"+Cstr(ii)+"</a></SPAN> "
			end if
			if ii=PageCount then exit for
		next
		if ii>p*10+10 then ii=ii-1
		response.write "</b>"
		if ii<PageCount then response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(ii+1)&URLPostfix&" title=下十页><font face=webdings>8</font></a></SPAN> "
		if CurPage=PageCount then
			response.write "<SPAN class=pagenumstatic><font face=webdings color="&Forum_body(8)&">:</font></SPAN>"
		else
			response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(PageCount)&URLPostfix&" title=尾页><font face=webdings>:</font></a></SPAN> "
		end if
	end if
end sub

Function GetBackground(UserVisual)
	dim VisualSplit
	if UserVisual="" then
		GetBackground=""
		Exit Function
	end if
	VisualSplit=Split(UserVisual,"|")
	if ubound(VisualSplit)<>24 then
		GetBackground=""
		Exit Function
	end if
	if VisualSplit(1)<>"" then
  	if localpic=1 then
			GetBackground="2/"&right("0000"&Split(VisualSplit(1),".")(0),4)&".Gif"
  	else
			GetBackground=right("00000000"&Split(VisualSplit(1),".")(0),8)&"/02/00/"
  	end if
	else
		GetBackground=""
	end if
End Function

Function hasASPPainter()
	Dim pic
	on error resume next
	Set pic = CreateObject("ASPPainter.Pictures.1")
	if err.Number=0 then
		on error goto 0
		hasASPPainter=true
	else
		err.clear
		on error goto 0
		hasASPPainter=false
	end if
End Function

Sub ShowVisualPhoto(bgWidth,bgHeight,PhotoBg,PhotoBodyBack,PhotoBodyFore,PhotoFg,PhotoName,PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection,PhotoDate,DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection,UserCount,UserVisualSplit,LayerNo,OuterLeft,OuterTop,OuterWidth,OuterHeight,InnerLeft,InnerTop,InnerWidth,InnerHeight,Direction,UserNameSplit,NameLeft,NameTop,NameFont,NameSize,NameBold,NameItalic,NameColor,NameDirection,isMask,isUpload,isPhotoMe)
	response.write "<DIV id=""ThePhoto"" style=""border:1px solid "&forum_body(27)&";left:0;overflow:hidden;width:"&bgWidth&";position:relative;top:0;height:"&bgHeight&""">"
	if isUpload then
		response.write "<img src="""&PhotoPath&"\"&photobg&""" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:1;display:"
	else
		response.write "<img src="""&PicPath&photobg&""" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:1;display:"
	end if
	if photobg<>"" then
		response.write "inline;"
	else
		response.write "none;"
	end if
	response.write """ name=""photo_bg"">"
	if bgWidth=280 then
		response.write "<img src="""&PicPath&photoBodyBack&""" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:2;display:"
		if photoBodyBack<>"" then
			response.write "inline;"
		else
			response.write "none;"
		end if
		response.write """ name=""photo_bodyback"">"
	end if
	for i=0 to 2
		response.write GetAccoutermentDiv(i,uservisualsplit(UserCount+i),i*2+3,OuterLeft(UserCount+i),OuterTop(UserCount+i),OuterWidth(UserCount+i),OuterHeight(UserCount+i),InnerLeft(UserCount+i),InnerTop(UserCount+i),InnerWidth(UserCount+i),InnerHeight(UserCount+i),Direction(UserCount+i))
	next
	for i=0 to usercount-1
		response.write GetBodyDiv(i,uservisualsplit(i),LayerNo(i),OuterLeft(i),OuterTop(i),OuterWidth(i),OuterHeight(i),InnerLeft(i),InnerTop(i),InnerWidth(i),InnerHeight(i),Direction(i),isPhotoMe)
	next
	for i=0 to usercount-1
		response.write GetNameDiv(i,usernamesplit(i),usercount*20+i*5,NameLeft(i),NameTop(i),NameFont(i),NameSize(i),NameBold(i),NameItalic(i),NameColor(i),NameDirection(i))
	next
	for i=0 to 2
		response.write GetAccoutermentDiv(i+3,uservisualsplit(UserCount+i+3),UserCount*25+i*2,OuterLeft(UserCount+i+3),OuterTop(UserCount+i+3),OuterWidth(UserCount+i+3),OuterHeight(UserCount+i+3),InnerLeft(UserCount+i+3),InnerTop(UserCount+i+3),InnerWidth(UserCount+i+3),InnerHeight(UserCount+i+3),Direction(UserCount+i+3))
	next
	if bgWidth=280 then
		response.write "<img src="""&PicPath&photoBodyFore&""" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:"&(usercount*25+10)&";display:"
		if photoBodyFore<>"" then
			response.write "inline;"
		else
			response.write "none;"
		end if
		response.write """ name=""photo_bodyfore"">"
	end if
	if bgWidth=280 then
		response.write "<img src="""&PicPath&photoFg&""" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:"&(usercount*25+11)&";display:"
		if photoFg<>"" then
			response.write "inline;"
		else
			response.write "none;"
		end if
		response.write """ name=""photo_fg"">"
	end if
	response.write GetPhotoNameDiv(photoname,usercount*25+15,PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection)
	response.write GetDateDiv(PhotoDate,usercount*25+20,DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection)
	if isMask then
		if bgWidth=140 then
			response.write "<img src=""images/img_visual/unfinished_s.gif"" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:"&(usercount*25+25)&";display:inline;"" name=""photo_mask"">"
		else
			response.write "<img src=""images/img_visual/unfinished.gif"" style=""padding:0;position:absolute;top:0;left:0;width:"&bgWidth&";height:"&bgHeight&";z-index:"&(usercount*25+25)&";display:inline;"" name=""photo_mask"">"
		end if
	end if
	response.write "</DIV>"
End Sub
%>