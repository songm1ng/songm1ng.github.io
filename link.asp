<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc.asp"-->
<!--#include file="Inc/function.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title>友情链接- <%=SiteTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
	<meta name="keywords" content="<%=Sitekeywords%>" />
	<meta name="description" content="<%=Sitedescription%>" />
	<%if SitePath="/" then%>
	<link href="<%=SitePath%><%=skin%>head.css" rel="stylesheet" type="text/css">
	<%else%>
	<link href="<%=SitePath%><%=skin%>header.css" rel="stylesheet" type="text/css">
	<%end if%>
	<link href="<%=SitePath%>skin/qqvideo/default.css" rel="stylesheet" type="text/css">
	<link href="<%=SitePath%>skin/qqvideo/index.css" rel="stylesheet" type="text/css">
</head>
<body>
	<div style="z-index:40000">
	  <%=head%>
		<div class="vtop">
		<div class="vtopbg">
		
		<div class="logo">
		<span class="tlogo"><a href="http://<%=SiteUrl%>" title="<%=SiteTitle%>"></a></span>
		<span class="toptips">身边遇到的问题解决方法分享.</span>
		</div>
		

		<div class="menus">
        <span class="tmenu">
        <a href="http://<%=SiteUrl%><%=SitePath%>bbs/" target="_blank">程序讨论▼</a>
        <a href=# onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('pcook.com.cn');">设为首页</a>
        <script language="JavaScript">
function bookmarkit()
 {window.external.addFavorite('http://pcook.com.cn','★ Pcook CMS 泡客 ');}
if (document.all)document.write('<a href="#" onClick="bookmarkit();" title="把“Pcook CMS 泡客”加入您的收藏夹！">收藏本站</a>')
</script>
        <a target="_blank" href="http://<%=SiteUrl%><%=SitePath%>sitemap.asp" class="end">网站地图</a>
        </span>
		</div>
		
		
		
		<div class="topmenu">
		
      <div class="menunr">
      <%=Menu%>
      </div>
		
		<div class="topmenubg">
      <span class="lbg"></span>
      <span class="mbg"></span>
      <span class="rbg"></span>
		</div>
		
		</div>

		</div>
	</div>
	<div class="tlinehead"></div>
		
<%
set rs4 = server.CreateObject ("adodb.recordset")
sql="select UserMoney from pcook_User where UserName='"& UserName &"'"
rs4.open sql,conn,1,1
mymoney=rs4("UserMoney")
rs4.close
set rs4=nothing
%>
	
	<div class="vmovie">
		<div class="vmoviebg">

<!-- 左侧开始 -->

		<div class="vleft">
		

		
<!-- 一个循环段落开始 -->

		<div class="movelist">
		
		<div class="lpic"></div>

		<div>
		<span class="mbg3">
		<span class="title">您现在的位置：<a href="Index.asp">首页</a> >> <a href="link.asp">所有链接</a> <a href="?ac=add" style="background:#FF0066;color:#ffffff;padding:3px 5px;font-weight:bold;">申请链接</a></span>
		<span class="more"></span>
		</span>
		</div>
		
		<div class="lpic2"></div>
		</div>

<%If request("ac")="" then%>	

		<div class="movies end">

		<div class="vmoviesa">

		<div class="movernrsa">
		<ul>
			
<%
set rs=server.createobject("adodb.recordset")
sql="select * from pcook_link where yn=1 order by id desc"	
rs.open sql,conn,1,1  
if rs.eof and rs.bof then
 response.Write("暂时没有链接")
else
dim currentpage 
maxperpage=15 '信息显示条数           
rs.pagesize=maxperpage              
currentpage=request.querystring("page")              
if currentpage="" then              
currentpage=1              
elseif currentpage<1 then              
currentpage=1              
else              
currentpage=clng(currentpage)              
	if currentpage > rs.pagecount then              
	currentpage=rs.pagecount              
	end if              
end if              
              
if not isnumeric(currentpage) then              
currentpage=1              
end if              
dim totalput,n              
totalput=rs.recordcount              
if totalput mod maxperpage=0 then              
n=totalput\maxperpage              
else              
n=totalput\maxperpage+1              
end if              
if n=0 then              
n=1              
end if              
rs.move(currentpage-1)*maxperpage              
i=0     
w=1
NoI=currentpage*maxperpage-maxperpage    
do while i< maxperpage and not rs.eof
NoI=NoI+1
%>	

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="gcontent" height="10"></td>
    </tr>
  <tr>
    <td width="40%" height="30" valign="middle" style="background-color:#fafafa;padding:5px 0 0 10px;"><font color=red><B><%=NoI%>.</B>　</font>网站名称：<font color=red><%=left(rs("UserName"),10)%>&nbsp; </font></td>
    <td height="30" valign="middle" style="background-color:#fafafa;padding:5px 0 0 10px;">　　网站地址：<a href=<%=rs("Title")%> target="_blank"><%=rs("Title")%></a></td>
    </tr>
  <tr>
    <td colspan="2" height="30" valign="middle" style="background-color:#fafafa;padding:0 0 5px 10px;">　　网站简介：<%=dvHTMLEncode(rs("Content"))%></td>
    </tr>
</table>

<%
  i=i+1      
  w=w+1
  rs.movenext
  loop
  end if
  rs.close
  set rs=nothing
%>
  
  
		</ul>
		</div>
		
		</div>

		</div>
		
<!-- 一个循环段落结束 -->


<div id="page">
<table border="0" cellspacing="5" cellpadding="2" align="center"><tr>
<%k=currentpage                                                                                              
   	if k<>1 then%>
<TD class="page_css_1"><a href="Guestbook.asp">首&nbsp;&nbsp;页</a></TD>
<TD class="page_css_1"><a href="?Page=<%=k-1%>">上一页</a></TD>
<%else%>
<TD class="page_css_1">首&nbsp;&nbsp;页</TD>
<TD class="page_css_1">上一页</TD>
<%end if%>

<%
	yw1=currentpage-2
	  if yw1<1 then 
	  	yw1=1 
	  end if
	yw2=currentpage+2
	  if yw2>n then 
	  	yw2=n
	  end if
  for i = yw1 to yw2
    if i = currentpage then%>
<TD class="page_css_2"><%=i%></TD>
<%else%>
<TD class="page_css_2"><a href="?Page=<%=i%>"><%=i%></a></TD>
<%
end if
next
%>

<% if k<>n then %>
<TD class="page_css_1"><a href="?Page=<%=k+1%>">下一页</a></TD>
<TD class="page_css_1"><a href="?Page=<%=n%>">尾&nbsp;&nbsp;页</a></TD>
<%else%>
<TD class="page_css_1">下一页</TD>
<TD class="page_css_1">尾　页</TD>
<%end if%>

<TD class="page_css_3">页次:<%=currentpage%>/<%=n%>页 共<font color="#009900"><b> <%=totalput%> </b></font>条记录 <%=maxperpage%>条/页</TD>
</tr></table>
</div>
	

<%elseif request("ac")="add" then%>

		<div class="movies end">

		<div class="vmoviesa">

		<div class="movernrsa">

<script language=javascript>
function chk()
{
	if(document.form.UserName.value == "" || document.form.UserName.value.length > 20)
	{
	alert("不能提交申请，网站名称为空或大于20字符！");
	document.form.UserName.focus();
	document.form.UserName.select();
	return false;
	}
	if(document.form.title.value == "" || document.form.title.value.length > 50)
	{
	alert("不能提交申请，网站地址为空或大于50个字符！");
	document.form.title.focus();
	document.form.title.select();
	return false;
	}
	if(document.form.content.value == "")
	{
	alert("请填写网站简介！");
	document.form.content.focus();
	document.form.content.select();
	return false;
	}
	if(document.form.code.value == "")
	{
	alert("请填写验证码！");
	document.form.code.focus();
	document.form.code.select();
	return false;
	}
return true;
}
</script>

<form onSubmit="return chk();" method="post" name="form" action="?ac=post">
  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#ffffff">
    <tr>
      <td height="30" class="adgs" align="right">网站名称：</td>
      <td height="15" class="adgs">
      <input name="UserName" type="text" id="UserName" maxlength="20" style="width:240px;border:1px solid #ccc;"> <font color=#ff0000>*</font></td>
    </tr>
    <tr>
      <td height="30" class="adgs" align="right">网站地址：</td>
      <td height="15" class="adgs">
      <input name="title" type="text" id="title" maxlength="50" style="width:240px;border:1px solid #ccc;"> <font color=#ff0000>*</font> 网址前请添加：http://</td>
    </tr>
    <tr>
      <td height="15" class="adgs" align="right" valign="top">网站简介：</td>
      <td height="15" class="adgs"><textarea name="content" cols="30" rows="6" id="content" style="width:400px;border:1px solid #ccc;margin:0;padding:0;height:200px;"></textarea> <font color=#ff0000>*</font></td>
    </tr>
	<tr>
      <td height="30" class="adgs" align="right">验证码：</td>
      <td height="15" class="adgs">
      <input name="code" type="text" id="code" size="8" maxlength="5" style="border:1px solid #ccc;"/>
        <img src="Inc/code.asp" border="0" alt="看不清楚请点击刷新验证码" style="cursor : pointer;" onClick="this.src='Inc/code.asp'"/>
	</td>
    </tr>
    <tr>
      <td height="30" align="center"></td>
      <td height="30" class="adgs" align="left" ><input type="submit" name="Submit2" class="borderall1" value=" 发 布 "> <Input name="Cancel" type=reset  class="borderall1" id=Reset value=" 重 填 "></td>
    </tr>
  </table>
</form>


		</div>
		
		</div>

		</div>
		
			<% elseif request("ac")="post" then
	dim UserName,Title,Content
	UserName = trim(request.form("UserName"))
	Title = trim(request.form("Title"))
	Content = request.form("Content")
	mycode = trim(request.form("code"))
	if mycode<>Session("getcode") then
	Response.Write("<script language=javascript>alert('请输入正确的认证码！');javascript:history.back();</script>") 
	Response.End 
	end if

  	if Instr(Content,"傻B")>0 or Instr(Content,"操你妈")>0 or Instr(Content,"胡锦")>0 or Instr(Content,"江泽")>0 or Instr(Content,"娘")>0 or Instr(Content,"[url")>0 or HasChinese(Content)=false then
	Response.Write "<script language=javascript>alert('请不要乱发信息,谢谢!');javascript:history.back();</script>"
	Response.End
	end if

	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from pcook_link"
	rs.open sql,conn,1,3
		if UserName="" or Title="" or Content="" then
			response.Redirect "link.asp?ac=add"
		end if

		rs.AddNew 
		rs("UserName")			=UserName
		rs("Title")				=Title
		rs("Content")			=Content
		rs("AddIP")				=Request.ServerVariables("REMOTE_ADDR")
		If linkoff=1 then
		rs("yn")				=0
		else
		rs("yn")				=1
		end if
		rs.update
		If linkoff=1 then
		Response.Write("<script language=javascript>alert('恭喜你,提交成功,但需要管理员审核后才能显示出来!');this.top.location.href='link.asp';</script>")
		else
		Response.Write("<script language=javascript>alert('恭喜你,提交成功');this.top.location.href='link.asp';</script>")
		end if
		rs.close
		Set rs = nothing
  end if %>
		</div>
		
<!-- 左侧结束 -->

<!-- 右侧开始 -->

		<div class="vright">
		<div>
		<div  id="upNotice" class="logout" >
		<div class="titles">
		
		<div class="down_con">
		<%Call ShowAD(1)%>	<!-- Pcook 下载 -->
		</div>
		
		</div>
		
<!-- 公告开始 -->
		<div class="nrlist">
		<div class="bbg2">
		<span class="cib3"></span>
		<span class="mbg2"></span>
		<span class="cib4"></span>
		</div>
		<div id="unlNotice" class="tbg5">
		<%If ""&ad3&""=1 then%><%Call ShowAD(4)%><%End If%>	<!-- Google 广告 -->
		</div>
		<div class="bbg2">
		<span class="cib"></span>
		<span class="mbg"></span>
		<span class="cib2"></span>
		</div>
		</div>
<!-- 公告结束 -->
		
		</div>

		</div>
		
		
		
		
		
		<div class="down_con">
		<%Call ShowAD(5)%>	<!-- 预留右侧广告 -->
		</div>




		
		<div class="hotmovie">
		<span>
		<span class="cit"></span>
		<span class="mbg"></span>
		<span class="cit2"></span>
		</span>
		<span class="tbg">
		<span class="title" style="color:red;">工作手记 原创</span>
		<span class="more"><a target="_blank" href="class.asp?ID=11">更多>></a></span>
		<span class="line"></span>
		
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 8 ID,Title,ClassID,TitleFontColor,DateAndTime from pcook_Article where ClassID=11 and yn = 0 order by DateAndTime desc,id desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err)

set rsClass=server.createobject("adodb.recordset")
sql = "select * from pcook_Class where ID="&rs1("ClassID")&""
rsClass.open sql,conn,1,1  
ClassName=rsClass("ClassName")
rsClass.close
set rsClass=nothing
%>
		
		<div class="movies hbg">
		<ul>
		<li class="names"><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank">·<%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),30)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),30)&"") end if%></a></li>
		</ul>
		</div>
		
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>
		</span>
		<span>
		<div class="cit3"></div>
		<div><span class="mbg2"></span></div>
		<div class="cit4"></div>
		</span>
		</div>
		
		
		<div class="vshow">
		<span>
		<span class="cit5"></span>
		<span class="mbg"></span>
		<span class="cit6"></span>
		</span>
		<span class="tbg">
		<span class="title" style="color:red;">站长推荐</span>
		<span class="more"><a href="hot.asp">更多>></a></span>
		<span class="line"></span>
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 5 ID,Title,Content,DateAndTime,Hits,TitleFontColor from pcook_Article where yn = 0 and IsHot = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>
		
		<div class="movies">
		<ul>
		<li  class="names" lang="5">·<a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank"><%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("Title")),25)&"</font>") else Response.Write(""&left(LoseHtml(rs1("Title")),25)&"") end if%></a></li>
		</ul>
		<span class="say">
		<p><%=left(trim(LoseHtml(rs1("Content"))),45)%></p>
		</span>
		</div>
		
<%
  rs1.movenext
  loop
  else
  Response.Write("没有")
  end if
  rs1.close
  set rs1=nothing
%>
		

		</span>
		
		<span>
		<div><span class="cit7"></span></div>
		<div><span class="mbg2"></span></div>
		<div><span class="cit8"></span></div>
		</span>
		</div>
		</div>
		</div>
	</div>

	<div class="vbottoms">
	<div class="vbottom">
	<p>
		<%Call ShowAD(6)%>　<%=SiteTcp%>	<!-- 网站底部 -->
	</p>
	<br style="line-height:3px;">
	</div>
</div>
	</div>
<SCRIPT LANGUAGE=JAVASCRIPT><!-- 
if (top.location != self.location)top.location=self.location;
// --></SCRIPT>
</body>
</html>
