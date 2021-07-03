<%
Function Head
	If useroff=1 then
	Response.Write("<table width=""980"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" >") & VbCrLf
	If Request.Cookies("pcook")("ID")="" then
      Response.Write("		<form action="""&SitePath&"User/Userlogin.asp?action=login"" method=""post"" name=loginForm>") & VbCrLf
      Response.Write("		<tr><td width=""170""><div align=""left"">用户名：<input name=""Username"" class=""borderall"" type=""text"" size=""15"" style=""width:100px;height:15px;""></div></td>") & VbCrLf
      Response.Write("		<td height=""30"" width=""170""><div align=""left"">密　码：<input name=""PassWord"" class=""borderall"" type=""password"" size=""15"" style=""width:100px;height:15px;""></div></td>") & VbCrLf
      Response.Write("		<td height=""30"" width=""160""><div align=""left""><input type=""submit"" name=""Submit"" value=""登录"" class=""borderall1"" style=""height:18px;""/>") & VbCrLf
      Response.Write("			<input type=""button"" name=""button"" value=""注册新用户"" class=""borderall1"" style=""height:18px;"" onClick=""window.location.href='"&SitePath&"User/userreg.asp'"" /></div></td>") & VbCrLf
      Response.Write("		<td height=""30""><div align=""right""><script type=""text/javascript"" src="""&Sitepath&"inc/date.js""></script></div></td></tr>") & VbCrLf
      Response.Write("		</form>") & VbCrLf
	else
	set rs4 = server.CreateObject ("adodb.recordset")
	sql="select * from pcook_User where id="& Request.Cookies("pcook")("ID") &""
	rs4.open sql,conn,1,1
	mymoney=rs4("UserMoney")
	dengjipic=rs4("dengjipic")
	dengji=rs4("dengji")
	rs4.close
	set rs4=nothing	
	Response.Write("		<tr><td width="""" height=""30""><div align=""left"">欢迎你："&Request.Cookies("pcook")("UserName")&"，"&moneyname&"：<b>"&mymoney&"</b>　等级："&dengji&"　<a href="""&SitePath&"User/UserAdd.asp?action=list"">我的文章</a>　<a href="""&SitePath&"User/UserAdd.asp?action=add"">发表文章</a>　<a href="""&SitePath&"User/UserAdd.asp?action=useredit"">修改资料</a>　<a href="""&SitePath&"User/UserLogin.asp?action=logout"">退出</a></div></td>") & VbCrLf
      Response.Write("		<td height=""30""><div align=""right""><script type=""text/javascript"" src="""&Sitepath&"inc/date.js""></script></div></td></tr>") & VbCrLf
	end if
	Response.Write("</table>") & VbCrLf
	else
	Response.Write("<div id=""clear""></div>") & VbCrLf
	End if
End function

Function Menu
	Response.Write("		<a href="""&SitePath&"index.asp"">首页</a><span></span>") & VbCrLf
	
set rs8=conn.execute("select * from pcook_class Where IsMenu=1 order by num asc")
do while not rs8.eof

	if seo=1 then
	Response.Write("<a href="""&SitePath&"class_"&rs8("ID")&".html"" title="""&rs8("ReadMe")&""" target="""&rs8("target")&""">"&rs8("className")&"</a><span></span>") & VbCrLf
	elseif seo=2 then
	Response.Write("<a href="""&SitePath&"class.asp?ID="&rs8("ID")&""" title="""&rs8("ReadMe")&""" target="""&rs8("target")&""">"&rs8("className")&"</a><span></span>") & VbCrLf
	elseif seo=3 then
	Response.Write("<a href="""&SitePath&"class.asp?ID="&rs8("ID")&""" title="""&rs8("ReadMe")&""" target="""&rs8("target")&""">"&rs8("className")&"</a><span></span>") & VbCrLf
	elseif seo=4 then
	Response.Write("<a href="""&SitePath&"class.asp?ID="&rs8("ID")&""" title="""&rs8("ReadMe")&""" target="""&rs8("target")&""">"&rs8("className")&"</a><span></span>") & VbCrLf
	end if
	
rs8.movenext
loop
	Response.Write("<a href="""&SitePath&"Guestbook.asp"" title=""用户留言"">留言簿</a><span></span>") & VbCrLf
	Response.Write("<a href="""&SitePath&"bbs/"" title=""程序讨论区"" class=""up"" target=""_blank""></a>") & VbCrLf
rs8.close
End function


Function search
	Response.Write("<div style=""float:left;margin-top:5px;margin-bottom:-13px;font-size:14px;color:#04A604;"">") & VbCrLf
	Response.Write("<form id=""form1"" name=""form1"" method=""get"" action="""&SitePath&"Search.asp"">") & VbCrLf
	Response.Write("<B>Search : </B><input name=""KeyWord"" type=""text"" id=""KeyWord"" maxlength=""10"" class=""borderall"" style=""width:140px;""/>") & VbCrLf
	Response.Write("<input type=""submit"" name=""Submit"" value=""搜 索"" class=""borderall1""/>") & VbCrLf
	Response.Write("</form>") & VbCrLf
	Response.Write("</div>") & VbCrLf
End Function
%>