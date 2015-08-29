<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_Tigerset.asp" -->
<%
call nav()
stats="老 虎 机"
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本社区老 虎 机的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
else
	%><bgsound ID="bgs" LOOP="-1">
	<script language="VBScript">
		dim j,k,l,i,h,m,n,o,Wins
		
		sub play()
			bgs.src="duchang_img/1.mid"
			if form1.e1.value+form1.e2.value+form1.e3.value+form1.e4.value+form1.e5.value+form1.e6.value+form1.e7.value=0 then
				msgbox"您没有选择下注",0
				exit sub
			end if
			
			ee1.disabled=true
			ee2.disabled=true
			ee3.disabled=true
			ee4.disabled=true
			ee5.disabled=true
			ee6.disabled=true
			ee7.disabled=true
			eeb.disabled=true
			eer.disabled=true
			
			randomize
			l=fix(72*rnd+48)'匀速转多少2~5圈
			randomize
			h=fix(144*rnd+0)
			if h<4 then
				h=13
			elseif h<12 then
				h=1
			elseif h<24 then
				h=int(2*rnd)
				if h<1 then
					h=4
				else
					h=16
				end if
			elseif h<40 then
				h=int(2*rnd)
				if h<1 then
					h=10
				else
					h=22
				end if
			elseif h<60 then
				h=int(3*rnd)
				if h<1 then
					h=0
				elseif h<2 then
					h=7
				else
					h=18
				end if
			elseif h<84 then
				h=int(3*rnd)
				if h<1 then
					h=8
				elseif h<2 then
					h=14
				else
					h=20
				end if
			elseif h<112 then
				h=int(5*rnd)
				if h<1 then
					h=2
				elseif h<2 then
					h=5
				elseif h<3 then
					h=11
				elseif h<4 then
					h=17
				else
					h=21
				end if
			else
				h=int(7*rnd)
				if h<1 then
					h=3
				elseif h<2 then
					h=6
				elseif h<3 then
					h=9
				elseif h<4 then
					h=12
				elseif h<5 then
					h=15
				elseif h<6 then
					h=19
				else
					h=23
				end if
			end if
			k=1
			i=0
			j=1
			toplay()
		end sub
		
		sub toplay()
			k=k+1	
			for x=1 to 24
				document.all("f"&x).bgcolor="#FFFFFF"
			next		
			document.all("f"&j).bgcolor="red"
			j=j+1
			if j>24 then j=1
			if k<=l-h then
				settimeout "toplay()",20
			else
				i=i+1
				if i>h then
					Wins=Win(j-1)
					if Win(j-1)>0 then
						form1.win.value=Win(j-1)
						bgs.src="duchang_img/3.mid"
						bgs.loop="0"
					else
						bgs.src="duchang_img/2.mid"
						bgs.loop="0"
					end if
					form1.submit()
					exit sub
				end if
				settimeout "toplay()",20*i
			end if
		end sub
		
		function Win(NN)
			if NN=3 or NN=6 or NN=9 or NN=12 or NN=15 or NN=19 or NN=23 then
				NN=document.form1.e1.value
			elseif NN=2 or NN=5 or NN=11 or NN=17 or NN=21 then
				NN=2*document.form1.e2.value
			elseif NN=8 or NN=14 or NN=20 then
				NN=4*document.form1.e3.value
			elseif NN=0 or NN=7 or NN=18 then
				NN=8*document.form1.e4.value
			elseif NN=10 or NN=22 then
				NN=10*document.form1.e5.value
			elseif NN=4 or NN=16 then
				NN=20*document.form1.e6.value
			elseif NN=13 then
				NN=40*document.form1.e7.value
			else
				NN=0
			end if
			Win=NN
		end function
		
		sub ResetE()
			document.form1.e1.value=0
			document.form1.e2.value=0
			document.form1.e3.value=0
			document.form1.e4.value=0
			document.form1.e5.value=0
			document.form1.e6.value=0
			document.form1.e7.value=0
			zero()
		end sub
		
		sub zero()
			for e=1 to 7
				document.all("d"&e).innerHTML=0
			next
		end sub
		
		sub bet1()
			If document.form1.e1.value<<%=BetLimit%> then
				document.form1.e1.value=document.form1.e1.value+<%=BetLimit\10%>
			End if
			d1.innerHTML=document.form1.e1.value
		end sub
		
		sub bet2()
			If document.form1.e2.value<<%=BetLimit%> then
				document.form1.e2.value=document.form1.e2.value+<%=BetLimit\10%>
			End if
			d2.innerHTML=document.form1.e2.value
		end sub
		
		sub bet3()
			If document.form1.e3.value<<%=BetLimit%> then
				document.form1.e3.value=document.form1.e3.value+<%=BetLimit\10%>
			End if
			d3.innerHTML=document.form1.e3.value
		end sub
	
		sub bet4()
			If document.form1.e4.value<<%=BetLimit%> then
				document.form1.e4.value=document.form1.e4.value+<%=BetLimit\10%>
			End if
			d4.innerHTML=document.form1.e4.value
		end sub
	
		sub bet5()
			If document.form1.e5.value<<%=BetLimit%> then
				document.form1.e5.value=document.form1.e5.value+<%=BetLimit\10%>
			End if
			d5.innerHTML=document.form1.e5.value
		end sub
		
		sub bet6()
			If document.form1.e6.value<<%=BetLimit%> then
				document.form1.e6.value=document.form1.e6.value+<%=BetLimit\10%>
			End if
			d6.innerHTML=document.form1.e6.value
		end sub
		
		sub bet7()
			If document.form1.e7.value<<%=BetLimit%> then
				document.form1.e7.value=document.form1.e7.value+<%=BetLimit\10%>
			End if
			d7.innerHTML=document.form1.e7.value
		end sub
	</script>
	<script language="JavaScript">
		function MM_preloadImages() { //v3.0
		  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
		    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
		    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
		}
		MM_preloadImages('duchang_img/1.mid','duchang_img/2.mid','duchang_img/3.mid')
	</script>

	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
		<tr><th valign=middle colspan=2 align=center height=25><b>老 虎 机</b></th></tr>
		<tr>
			<td valign=middle class=tablebody1 height=100><CENTER>
				<table cellpadding=0 cellspacing=0 border=0 align=center>
					<tr>
						<td>
							<table cellpadding=3 cellspacing=1 border=0 width=100%>
								<tr>
									<td>
										<form name="form1" method="post" action="z_TigerRun.asp" target="run"><tr>
											<table border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" align="center">
												<tr bgcolor="#FFFFFF" align="center"> 
													<td height="50" width="50" id="f10"><img src="duchang_img/t5.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f11"><img src="duchang_img/t2.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f12"><img src="duchang_img/t1.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f13"><img src="duchang_img/t7.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f14"><img src="duchang_img/t3.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f15"><img src="duchang_img/t1.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f16"><img src="duchang_img/t6.gif" width="32" height="32"></td>
												</tr>
												<tr bgcolor="#FFFFFF" align="center"> 
													<td height="50" width="50" id="f9"><img src="duchang_img/t1.gif" width="32" height="32"></td>
													<td colspan="5" rowspan="5"><iframe src="z_TigerMain.asp" name="run" id=" " scrolling="no" width="254" height="254" marginwidth="0" marginheight="0" frameborder="0" align="default"></iframe></td>
													<td height="50" width="50" id="f17"><img src="duchang_img/t2.gif" width="32" height="32"></td>
												</tr>
												<tr bgcolor="#FFFFFF" align="center"> 
													<td height="50" width="50" id="f8"><img src="duchang_img/t3.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f18"><img src="duchang_img/t4.gif" width="32" height="32"></td>
												</tr>
												<tr bgcolor="#FFFFFF" align="center"> 
													<td height="50" width="50" id="f7"><img src="duchang_img/t4.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f19"><img src="duchang_img/t1.gif" width="32" height="32"></td>
												</tr>
												<tr bgcolor="#FFFFFF" align="center"> 
													<td height="50" width="50" id="f6"><img src="duchang_img/t1.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f20"><img src="duchang_img/t3.gif" width="32" height="32"></td>
												</tr>
												<tr bgcolor="#FFFFFF" align="center">
													<td height="50" width="50" id="f5"><img src="duchang_img/t2.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f21"><img src="duchang_img/t2.gif" width="32" height="32"></td>
												</tr>
												<tr bgcolor="#FFFFFF" align="center"> 
													<td height="50" width="50" id="f4"><img src="duchang_img/t6.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f3"><img src="duchang_img/t1.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f2"><img src="duchang_img/t2.gif" width="32" height="32"></td>
													<td height="50" width="50" id=f1><B>Lose</B></td>
													<td height="50" width="50" id="f24"><img src="duchang_img/t4.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f23"><img src="duchang_img/t1.gif" width="32" height="32"></td>
													<td height="50" width="50" id="f22"><img src="duchang_img/t5.gif" width="32" height="32"></td>
												</tr>
											</table>
											<table width="358" border="0" cellspacing="0" cellpadding="0" align="center">
												<tr> 
													<td> 
														<table border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" align="center">
															<tr bgcolor="#99CCFF" align="center"> 
																<td width="50" height="30" id=d7>0</td>
																<td width="50" height="30" id=d6>0</td>
																<td width="50" height="30" id=d5>0</td>
																<td width="50" height="30" id=d4>0</td>
																<td width="50" height="30" id=d3>0</td>
																<td width="50" height="30" id=d2>0</td>
																<td width="50" height="30" id=d1>0</td>
															</tr>
															<tr bgcolor="#99CCFF" align="center"> 
																<td width="50" height="50" id=ee7 onClick="bet7()" style="cursor:hand" title="下注" onMouseover="ee7.bgcolor='#75BAFF'" onMouseout="ee7.bgcolor='#99CCFF'"><img src="duchang_img/t7.gif" width="32" height="32"><br>x40</td>
																<td width="50" height="50" id=ee6 onClick="bet6()" style="cursor:hand" title="下注" onMouseover="ee6.bgcolor='#75BAFF'" onMouseout="ee6.bgcolor='#99CCFF'"><img src="duchang_img/t6.gif" width="32" height="32"><br>x20</td>
																<td width="50" height="50" id=ee5 onClick="bet5()" style="cursor:hand" title="下注" onMouseover="ee5.bgcolor='#75BAFF'" onMouseout="ee5.bgcolor='#99CCFF'"><img src="duchang_img/t5.gif" width="32" height="32"><br>x10</td>
																<td width="50" height="50" id=ee4 onClick="bet4()" style="cursor:hand" title="下注" onMouseover="ee4.bgcolor='#75BAFF'" onMouseout="ee4.bgcolor='#99CCFF'"><img src="duchang_img/t4.gif" width="32" height="32"><br>x8</td>
																<td width="50" height="50" id=ee3 onClick="bet3()" style="cursor:hand" title="下注" onMouseover="ee3.bgcolor='#75BAFF'" onMouseout="ee3.bgcolor='#99CCFF'"><img src="duchang_img/t3.gif" width="32" height="32"><br>x4</td>
																<td width="50" height="50" id=ee2 onClick="bet2()" style="cursor:hand" title="下注" onMouseover="ee2.bgcolor='#75BAFF'" onMouseout="ee2.bgcolor='#99CCFF'"><img src="duchang_img/t2.gif" width="32" height="32"><br>x2</td>
																<td width="50" height="50" id=ee1 onClick="bet1()" style="cursor:hand" title="下注" onMouseover="ee1.bgcolor='#75BAFF'" onMouseout="ee1.bgcolor='#99CCFF'"><img src="duchang_img/t1.gif" width="32" height="32"><br>x1</td>
															</tr>
														</table>
													</td>
												</tr>
												<tr> 
													<td> 
														<table width="358" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" align="center" height="30" STYLE="CURSOR:HAND">
															<tr> 
																<td id=eeb onClick="play()" bgcolor="#CCFF00" align="center" width="50%" title="买定离手" onMouseover="eeb.bgcolor='#DDDDDD'" onMouseout="eeb.bgcolor='#CCFF00'">买定离手</td>
																<td id=eer onClick="ResetE()" bgcolor="#CCFF00" align="center" width="50%" title="重新下注" onMouseover="eer.bgcolor='#DDDDDD'" onMouseout="eer.bgcolor='#CCFF00'">重新下注</td>
															</tr>
														</table>
													</td>
												</tr>
											</table>
											<input type="hidden" name="e7" value="0">
											<input type="hidden" name="e6" value="0">
											<input type="hidden" name="e5" value="0">
											<input type="hidden" name="e4" value="0">
											<input type="hidden" name="e3" value="0">
											<input type="hidden" name="e2" value="0">
											<input type="hidden" name="e1" value="0">
											<input type="hidden" name="win" value="0">
										</form>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</CENTER></td>
		</tr>
	</table>
<%end if
call footer()
%>