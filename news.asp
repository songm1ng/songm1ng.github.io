<!--#include file="Inc/conn.asp"-->
<!--#include file="Inc/Function_Page.asp"-->
<!--#include file="Inc.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title>������������б� - <%=SiteTitle%></title>
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
</head>
<body><%call spiderbot()%>
	<div style="z-index:40000">
	  <%=head%>
		<div class="vtop">
		<div class="vtopbg">
		
		<div class="logo">
		<span class="tlogo"><a href="http://<%=SiteUrl%>/" title="<%=SiteTitle%>"></a></span>
		<span class="toptips">������������������������.</span>
		</div>
		

		<div class="menus">
        <span class="tmenu">
        <a href="http://<%=SiteUrl%><%=SitePath%>bbs/" target="_blank">�������ۨ�</a>
        <a href=# onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('pcook.com.cn');">��Ϊ��ҳ</a>
        <script language="JavaScript">
function bookmarkit()
 {window.external.addFavorite('http://pcook.com.cn','�� Pcook CMS �ݿ� ');}
if (document.all)document.write('<a href="#" onClick="bookmarkit();" title="�ѡ�Pcook CMS �ݿ͡����������ղؼУ�">�ղر�վ</a>')
</script>
        <a target="_blank" href="http://<%=SiteUrl%><%=SitePath%>sitemap.asp" class="end">��վ��ͼ</a>
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

<!-- ��࿪ʼ -->

		<div class="vleft">
		

		
<!-- һ��ѭ�����俪ʼ -->

		<div class="movelist">
		<div class="lpic"></div>
		
	
		<div>
		<span class="mbg3">
		<span class="title">�����ڵ�λ�ã�<a href="Index.asp">��ҳ</a> >> ��ʱ������������� >> �б�</span>
		<span class="more"></span>
		</span>
		</div>
		
		
		<div class="lpic2"></div>
		</div>
		<div class="movies ">
		
		<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 2 ID,Title,Images,Content,TitleFontColor from  pcook_Article where yn = 0 and Images<>'' order by DateAndTime desc,ID desc"
rs1.open sql1,conn,1,3
NoI=0
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
NoI=NoI+1
%>  
		
		<div class="vmovies">
		
		<div class="movie">
		<a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>">
		<div class="icon"><span></span></div>
		<div class="imgs"><img src="<%=SitePath%><%=SiteUp%>/<%=rs1("Images")%>" alt="" /></div>
		</a>
		</div>
		
		<div class="movernr">
		<ul>
		<li class="titles"><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>"><%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),13)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),13)&"") end if%></a></li>
		<li><%=left(LoseHtml(rs1("Content")),50)%></li>
		</ul>
		</div>
		
		</div>
		
  <%
  rs1.movenext
  loop
  else
  Response.Write("û��ͼƬ����")
  end if
  rs1.close
  set rs1=nothing
  %>
		

		</div>
		<div class="movies end">
		
<%
Set mypage=new xdownpage
mypage.getconn=conn
mypage.getsql="select ID,Title,DateAndTime,Hits,dig,IsTop,Images,TitleFontColor,Content from pcook_Article Where yn=0 order by IsTop desc,ID desc"
mypage.pagesize=""&artlistnum&""
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>

		<div class="vmoviesa">

		<div class="movernrsa">
		<ul>
		<li class="titlesa">  <span class="note"><!--<span id="dig<%=Rs("id")%>"><a href="javascript:Dig(<%=Rs("id")%>);" style="color:#ff6600;font-size:12px;font-weight:bold;">��һ��</a></span> �Ѷ� <span id="<%=Rs("id")%>"><%=Rs("dig")%></span> �Ρ�-->�����<font color="cc0000"><B><%=rs("Hits")%></B></font></a> �Ρ�<%=FormatDate(rs("DateAndTime"),5)%></span>��&nbsp;<%If rs("IsTop")=1 then Response.Write("<font color=red size=2>[�̶�] </font>") end if%><%If rs("Images")<>"" then Response.Write("<font color=blue>[ͼ]</font>") end if%><a href="<%If seo=1 then%>Article_<%=rs("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs("ID")%><%elseif seo=4 then%>Article/<%=rs("ID")%>.html<%end if%>" target="_blank"><B><%if rs("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs("TitleFontColor")&""">"&rs("Title")&"</font>") else Response.Write(""&rs("Title")&"") end if%></B></a> </li>
				<%If artlist=0 then%><li class="titlesb">����<%=left(LoseHtml(rs("Content")),100)%> <%End if%>  </li>
		</ul>
		</div>
		
		</div>
		
<%
        rs.movenext
    else
         exit for
    end if
next
%>



		</div>
		
<!-- һ��ѭ��������� -->

	<div class="movies end">
<table border="0" cellspacing="3" cellpadding="2" align="center"><tr><%=mypage.showpage()%></tr></table>
  </div>
		</div>
		
<!-- ������ -->

<!-- �Ҳ࿪ʼ -->

		<div class="vright">
		<div>
		<div  id="upNotice" class="logout" >
		<div class="titles">
		
		<div class="down_con">
		<%Call ShowAD(1)%>	<!-- Pcook ���� -->
		</div>
		
		</div>
		
<!-- ���濪ʼ -->
		<div class="nrlist">
		<div class="bbg2">
		<span class="cib3"></span>
		<span class="mbg2"></span>
		<span class="cib4"></span>
		</div>
		<div id="unlNotice" class="tbg5">
		<ul>
		<%If ""&ad3&""=1 then%><%Call ShowAD(4)%><%End If%>	<!-- Google ��� -->
		</div>
		<div class="bbg2">
		<span class="cib"></span>
		<span class="mbg"></span>
		<span class="cib2"></span>
		</div>
		</div>
<!-- ������� -->
		
		</div>

		</div>
		
		
		
		
		
		<div class="down_con">
		<%Call ShowAD(5)%>	<!-- Ԥ���Ҳ��� -->
		</div>




		
		<div class="hotmovie">
		<span>
		<span class="cit"></span>
		<span class="mbg"></span>
		<span class="cit2"></span>
		</span>
		<span class="tbg">
		<span class="title" style="color:red;">�����ּ� ԭ��</span>
		<span class="more"><a target="_blank" href="class.asp?ID=11">����>></a></span>
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
		<li class="names"><a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank">��<%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("title")),30)&"</font>") else Response.Write(""&left(LoseHtml(rs1("title")),30)&"") end if%></a></li>
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
		<span class="title" style="color:red;">վ���Ƽ�</span>
		<span class="more"><a href="hot.asp">����>></a></span>
		<span class="line"></span>
		
<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 3 ID,Title,Content,DateAndTime,Hits,TitleFontColor from pcook_Article where yn = 0 and IsHot = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>
		
		<div class="movies">
		<ul>
		<li  class="names" lang="5">��<a href="<%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>List.asp?ID=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<%end if%>" title="<%=rs1("title")%>" target="_blank"><%if rs1("TitleFontColor")<>"" then Response.Write("<font style=""color:"&rs1("TitleFontColor")&""">"&left(LoseHtml(rs1("Title")),25)&"</font>") else Response.Write(""&left(LoseHtml(rs1("Title")),25)&"") end if%></a></li>
		</ul>
		<span class="say">
		<p><%=left(trim(LoseHtml(rs1("Content"))),45)%></p>
		</span>
		</div>
		
<%
  rs1.movenext
  loop
  else
  Response.Write("û��")
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
		<%Call ShowAD(6)%>��<%=SiteTcp%>	<!-- ��վ�ײ� -->
	</div>
</div>
	</div>
<SCRIPT LANGUAGE=JAVASCRIPT><!-- 
if (top.location != self.location)top.location=self.location;
// --></SCRIPT>
<%
rsclass.close
set rsclass=nothing
%>
</body>
</html>