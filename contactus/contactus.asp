<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/login.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_login_STRING
Recordset1_cmd.CommandText = "SELECT * FROM contactus" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/ProjectWeb.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>联系我们</title>
<link href="contactus.css" rel="stylesheet" type="text/css">
<!-- InstanceEndEditable -->
<style type="text/css">
body {
	background-color: #D5DADE;
}
</style>
<link href="../main.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
</head>

<body>
<div class="webbody">
  <div class="head">
    <p><strong><img src="../indeximages/ground.jpg" width="900" height="204"></strong></p>
    <div class="navbar">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <th width="128" height="50" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../index.asp" class="navword">首页</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchteam/researchteamprofessor.asp" class="navword">团队成员</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchdirection/researchdirection.asp" class="navword">研究方向</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchresult/researchresult.asp" class="navword">论文成果</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../teamactivity/teamactivity.asp" class="navword">团队动态</a></th>
          <th scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="contactus.asp" class="navword">联系我们</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="../login/backsystemlogin.asp" class="navword">管理入口</a></th>
        </tr>
      </table>
      <br></div>  
      <div style="height:50">
    <p></p>
    </div>
    <!-- InstanceBeginEditable name="EditRegion" -->
    <div class="leaderdiv">
<table width="220" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th colspan="2" scope="col" class="leadertable" align="left" valign="middle"><img src="../indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle">&nbsp;&nbsp;学术带头人</th>
  </tr>
  <tr>
  <td height="5" colspan="2">
  </td>
  </tr>
  <tr>
    <td rowspan="3"  class="leaderphoto"><img src="../indeximages/教授.jpg" width="100" height="120"></td>
    <td  class="leaderrank">张三教授</td>
  </tr>
  <tr>
    <td class="leaderrank">机械电子系系主任</td>
  </tr>
  <tr>
    <td  class="leaderrank">智能机器人所所长</td>
  </tr>
    <tr>
  <td height="5" colspan="2">
  </td>
  </tr>
  <tr>
    <td colspan="2" class="leaderintroductionword"> &nbsp; 黄继光（1931年1月8日—1952年10月19日），民族英雄。1931年生于四川省中江县，中国人民志愿军第45师135团9连的通讯员。1952年10月19日在朝鲜上甘岭地区597.9高地牺牲,年仅21岁。被中国人民志愿军领导机关追记特等功，并授予“特级英雄”称号[1]  ；所在部队党委追授他为中国共产党正式党员；朝鲜民主主义人民共和国最高人民会议常务委员会授予他“朝鲜民主主义人民共和国英雄”称号和金星奖章和一级国旗勋章。</td>
  </tr>
</table>
</div>

        <div id="currentplace"><table width="230" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th class="currentptable" scope="col"><img src="../indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle"></th>
	<th class="currentptable" scope="col">当前位置：</th>
    <th class="currentptable" scope="col"><a href="../index.asp" class="dellink" 
onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">首页</a>-><a href="contactus.asp" class="dellink" 
onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">联系我们</a></th>
  </tr>
</table>
</div>
<div id="condivindex">
  <table border="0" width="570">
    <tr>
      <td width="30"  colspan="3"><img src="../indeximages/schoolplace.png" width="665" height="300"></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td  class="footsong">联系人：<%=(Recordset1.Fields.Item("conname").Value)%></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td  class="footsong">电话：<%=(Recordset1.Fields.Item("tel").Value)%></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td  class="footsong">邮箱：<%=(Recordset1.Fields.Item("email").Value)%></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td  class="footsong">邮编：<%=(Recordset1.Fields.Item("postcode").Value)%></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td  class="footsong">学校地址：<%=(Recordset1.Fields.Item("address").Value)%></td>
    </tr>
    <tr>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td  class="footsong">办公室地址：<%=(Recordset1.Fields.Item("detailaddress").Value)%></td>
    </tr>
  </table>

</div>
     <div class="footer"><table width="570" border="0" cellspacing="5" cellpadding="2" align="center">
  <tr>
    <th  align="center" id="footerwords">&copy;&nbsp;Copyright&nbsp;2007-2014&nbsp;百度百度百度课题组版权所有

</th>
  </tr>
  <tr>
    <td align="center" id="footerwords">北京市西长安街174号中南海新华门|邮编：100017</td>
  </tr>
  <tr>
    <td align="center" id="footerwords">电话：010-11111111|邮箱：xiaobai@163.com|推荐兼容模式打开</td>
  </tr>
</table>
</div>

	
	<!-- InstanceEndEditable --></div>

  
    
   
  
</div>
</body>
<!-- InstanceEnd --></html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
