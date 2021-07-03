<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title><%=SiteTitle%> 享受分享的乐趣，做更好的女人.</title>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
	<meta name="keywords" content="<%=Sitekeywords%>" />
	<meta name="description" content="<%=Sitedescription%>" />
	<%if SitePath="/" then%>
	<link href="<%=SitePath%><%=skin%>head.css" rel="stylesheet" type="text/css">
	<%else%>
	<link href="<%=SitePath%><%=skin%>header.css" rel="stylesheet" type="text/css">
	<%end if%>
	<link href="<%=SitePath%><%=skin%>default.css" rel="stylesheet" type="text/css">
	<link href="<%=SitePath%><%=skin%>index.css" rel="stylesheet" type="text/css">
  <script type="text/javascript" src="<%=SitePath%>inc/main.asp"></script>
  <script type="text/javascript" src="<%=SitePath%>Dig.asp"></script>
  <style type="text/css">
<!--
.STYLE1 {font-size: 14px}
-->
  </style>
</head>
<body><%call spiderbot()%>
	<div style="z-index:40000">
	  <%=head%>
		<div class="vtop">
		<div class="vtopbg">
		
		<div class="logo">
		<span class="tlogo"><a href="http://<%=SiteUrl%>/" title="<%=SiteTitle%>"></a></span>
		<span class="toptips STYLE1">女性资讯www.rearea.com.cn</span>		</div>
		

		<div class="menus">
        <span class="tmenu">
        <a href="http://<%=SiteUrl%><%=SitePath%>bbs/" target="_blank">热啊图吧▼</a>
        <a href=# onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('www.rearea.cn');">设为首页</a>
        <script language="JavaScript">
function bookmarkit()
 {window.external.addFavorite('http://www.rearea.com.cn','★ 热啊女性网 ');}
if (document.all)document.write('<a href="#" onClick="bookmarkit();" title="把“热啊女性网”加入您的收藏夹！">收藏本站</a>')
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
	

	
	<div class="vmovie">
		<div class="vmoviebg">
		  <%Call ShowAD(13)%><!-- 谷歌广告 -->		
<!-- 左侧开始 -->

		<div class="vleft">
		
<!-- FLASH 幻灯片开始 -->

		<div id="flashbg" class="flash">
		<script type="text/javascript">
				var swf_width=250;
				var swf_height=200;
				var config='5|0x000000|0xEEF3F6|80|0xffffff|0xffcc00|0x333333';
				//-- config 参数设置 -- 自动播放时间(秒)|文字颜色|文字背景色|文字背景透明度|按键数字颜色|当前按键颜色|普通按键色彩 --
				var files='',links='', texts='';<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 6 ID,Title,Images from pcook_Article where yn = 0 and Images<>'' order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>
				files+='|<%=SiteUp%><%If AspJpeg=1 Then %>/s250/<%Else%>/<%end if%><%=rs1("Images")%>';links+='|<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>';texts+='|<%=rs1("Title")%>'; 
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
  %>
				files=files.substring(1);links=links.substring(1);texts=texts.substring(1);
				document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
				document.write('<param name="movie" value="<%=skin%>swfnews.swf" />');
				document.write('<param name="quality" value="high" />');
				document.write('<param name="menu" value="false" />');
				document.write('<param name=wmode value="opaque" />');
				document.write('<param name="FlashVars" value="config='+config+'&bcastr_flie='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'" />');
				document.write('<embed src="<%=skin%>swfnews.swf" wmode="opaque" FlashVars="config='+config+'&bcastr_flie='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
				document.write('</object>');
				</script>
		</div>
		
<!-- FLASH 幻灯片结束 -->

<!-- 头条固定信息开始 -->
		
		<div class="flashnr">
		<div class="line"></div>
		<div class="cir"></div>
		<div class="mbg">
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 1 ID,Title,Content,author,CopyFrom,hits,DateAndTime,dig from pcook_Article where yn = 0 and IsTop = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>

		<ul>
		<li class="title"><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank" style="font-size:14px;color:#ff0000;"><%=left(LoseHtml(rs1("title")),40)%></a></li>
		<li class="title2">　　<%=left(LoseHtml(rs1("Content")),130)%>…</li>
		<li class="titless">
		作者：<span><%=rs1("Author")%></span>
		来源：<span><%=rs1("CopyFrom")%></span>
		点击：<span><%=rs1("Hits")%></span>
		被顶：<span id="<%=Rs1("id")%>"><%=Rs1("dig")%></span> 次 </li>
		</ul>

		<p> <a href="javascript:Dig(<%=Rs1("id")%>);">顶一下</a> Time : <%=rs1("DateAndTime")%></p>
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>				
		
		</div>
		<div class="bline"></div>
		<div class="cir2"></div>
		</div>
		
<!-- 头条固定信息结束 -->

		<%Call Label(5)%>

		<div class="tjmovies">
		<span class="flpic"></span>
		<span class="mbg">
		<div class="fbg">
		<%Call ShowAD(2)%>	<!-- 谷歌广告 -->
		<%Call ShowAD(3)%> <!-- 首页第一个横幅 -->
		</div>
		</span>
		<span class="flpic2"></span>
		</div>
		
<!-- 一个循环段落开始 -->

		<div class="movelist">
		<div class="lpic"></div>
		<div>
		<span class="mbg3">
		<span class="title">最近更新</span>
		<span class="more"><a href="news.asp" style="color:#ff3300;">点击更多</a></span>
		</span>
		</div>
		<div class="lpic2"></div>
		</div>

		<div class="movies end">
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 10 ID,Title,Content,ClassID,TitleFontColor,DateAndTime from pcook_Article where yn = 0 order by DateAndTime desc,id desc"
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

		<div class="vmovies">

		<div class="movernrs">
		<ul>
		<li class="titles"><span class="note"><%If ""&DateDiff("d",""&FormatDate(rs1("DateAndTime"),5)&"",date())&""=0 then Response.Write("<font color=FF0000>"&FormatDate(rs1("DateAndTime"),5)&"</FONT>") else Response.Write(""&FormatDate(rs1("DateAndTime"),5)&"") end if%></span><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank"><font size=1>·</font><%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),20)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),20)&"") end if%></a></li>
		</ul>
		</div>
		
		</div>
		
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>
		</div>
		
<!-- 一个循环段落结束 -->

<!-- 栏目调用开始 -->
<%
Dim rs,rs1
NoI=0
set rs=conn.execute("select * from pcook_Class Where IsIndex = 1 And link=0 order by num asc")
do while not rs.eof
NoI=NoI+1
Response.Write("		<div class=""movelist""><div class=""lpic""></div><div><span class=""mbg3"">") & VbCrLf
Response.Write("		<span class=""title"">"&rs("ClassName")&"</span>") & VbCrLf
Response.Write("		<span class=""more""><a href=""Class.asp?ID="&rs("ID")&""">更多>></a></span>") & VbCrLf
Response.Write("		</span></div><div class=""lpic2""></div></div><div class=""movies"">") & VbCrLf
If indeximg=1 then
%>
  <%
  set rs1=server.createobject("ADODB.Recordset")
  sql1="select Top 2 ID,Title,Images,Content,TitleFontColor from  pcook_Article where yn = 0 and Images<>'' and ClassID="&rs("ID")&" order by DateAndTime desc,ID desc"
  rs1.open sql1,conn,1,3
  NoI=0
  If Not rs1.Eof Then 
  do while not (rs1.eof or err) 
  NoI=NoI+1
  %>
    <div class="vmovies"><div class="movie">
      <a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>">
    <div class="icon"><span></span></div>
		<div class="imgs"><img src="<%=SitePath%><%=SiteUp%><%If AspJpeg=1 Then %>/s120/<%Else%>/<%end if%><%=rs1("Images")%>" alt="" /></div></a></div>
		
		<div class="movernr"><ul>
		<li class="titles"><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>"><%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),13)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),13)&"") end if%></a></li>
		<li><%=left(LoseHtml(rs1("Content")),50)%></li>
		</ul></div></div>
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>
<%End if
Response.Write("		<div id=""clear""></div>") & VbCrLf
Response.Write("		</div>") & VbCrLf
Response.Write("		<div class=""movies end"">") & VbCrLf
%>

<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top "&rs("IndexNum")&" ID,Title,Content,ClassID,TitleFontColor,DateAndTime from pcook_Article where ClassID="&rs("ID")&" and yn = 0 order by DateAndTime desc,id desc"
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

		<div class="vmovies"><div class="movernrs"><ul>
		<li class="titles"><span class="note"><%If ""&DateDiff("d",""&FormatDate(rs1("DateAndTime"),5)&"",date())&""=0 then Response.Write("<font color=FF0000>"&FormatDate(rs1("DateAndTime"),5)&"</FONT>") else Response.Write(""&FormatDate(rs1("DateAndTime"),5)&"") end if%></span><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank"><font size=1>·</font><%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),20)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),20)&"") end if%></a></li>
		</ul></div></div>
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>

<%
Response.Write("		</div>") & VbCrLf
rs.movenext
loop
rs.close
%>

<!-- 栏目调用结束 -->		
		
		</div>
		
<!-- 左侧结束 -->

<!-- 右侧开始 -->

		<div class="vright"><div>
		
<!-- Pcook下载和公告开始 -->
		<div  id="upNotice" class="logout" >
		<div class="titles"><div class="down_con"><%Call ShowAD(1)%>	<!-- Pcook 下载 --></div></div>	
		<div class="nrlist"><div class="bbg2">
		<span class="cib3"></span><span class="mbg2"></span><span class="cib4"></span></div>
		<div id="unlNotice" class="tbg2"><%Call Label(4)%> <!-- 公告调用 --></div>
		<div class="bbg2"><span class="cib"></span><span class="mbg"></span><span class="cib2"></span></div></div></div>
<!-- Pcook下载和公告结束 -->

<!-- 搜索开始 -->
		<div  id="upNotice" class="logout" >
		<div class="nrlist" style="margin-top:5px;">
		<div class="bbg2">
		<span class="cib3"></span><span class="mbg2"></span><span class="cib4"></span>
		</div>
		<div id="unlNotice" class="tbg2">
		<%=search%> <!--  搜索调用 修改在inc.asp文件中 -->
		</div>
		<div class="bbg2">
		<span class="cib"></span><span class="mbg"></span><span class="cib2"></span>
		</div></div></div>
<!-- 搜索结束 -->


		</div>	
		
		<div class="down_con">
		<%Call ShowAD(5)%>	<!-- 侧边预留广告 -->
		</div>
		
		<div class="hotmovie">
		<span>
		<span class="cit"></span>
		<span class="mbg"></span>
		<span class="cit2"></span>
		</span>
		<span class="tbg">
		<span class="title" style="color:red;">热门文章列表</span>
		<span class="more"><a target="_blank" href="hot.asp">更多>></a></span>
		<span class="line"></span>		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 10 ID,Title,Content,ClassID,TitleFontColor,DateAndTime from pcook_Article where yn = 0 and IsHot = 1  order by hits desc,id desc"
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
		<span class="more"><a target="_blank" href="hot.asp">更多>></a></span>
		<span class="line"></span>
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 8 ID,Title,Content,DateAndTime,Hits,TitleFontColor from pcook_Article where yn = 0 and IsHot = 1 order by ID desc"
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

			<div class="borderall">
  <div class="link"><span class="note"><a href="link.asp?ac=add"><font color=#333333>申请加入</font></a> <a href="link.asp"><font color=#333333>所有链接</font></a> </span><B style="font-size:14px;color:ff0000;">友情链接：</B> <a href="http://www.webmasterhome.cn/" target="_blank"><img src="<%=SitePath%><%=skin%>pr.gif" style="width:80px;height:15px;border:0px;" alt="PageRank" align=absmiddle></a> <font color=#2e90d7>不接受内容空洞、含有非法信息的网站，请在申请通过后24小时内在首页添加本站链接，7天内未添加本站链接亦予删除！</font>
<BR><BR style="line-height:5px;">
<div style="border-top:#ddd 1px solid;padding:5px;">
<ul>
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 27 * from pcook_link where yn = 1 order by ID asc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>
			<li style="float:left;text-align:left;width:11%;line-height:20px;padding:0 0 5px 0;">&nbsp;<a href="<%=rs1("Title")%>" title="<%=dvHTMLEncode(rs1("Content"))%>" target="_blank"><%=rs1("UserName")%></a></li>
  <%
  rs1.movenext
  loop
  else
  Response.Write("		<li>还没有链接!</li>")
  end if
  rs1.close
  set rs1=nothing
%>
</ul></div>
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
