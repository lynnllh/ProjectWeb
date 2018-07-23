<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/login.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_login_STRING
Recordset1_cmd.CommandText = "SELECT * FROM login" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "../backsystemindex.asp"
  MM_redirectLoginFailed = "backsystemlogin.asp"

  MM_loginSQL = "SELECT username, password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM login WHERE username = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_login_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 255, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 255, Request.Form("password")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/ProjectWeb.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>用户登录</title>
<link href="login.css" rel="stylesheet" type="text/css">
<!-- InstanceEndEditable -->
<style type="text/css">
body {
	background-color: #D5DADE;
}
</style>
<link href="../main.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" -->
<script type="text/javascript">
function MM_validateForm() { //v4.0
  if (document.getElementById){
    var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
    for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=document.getElementById(args[i]);
      if (val) { nm=val.name; if ((val=val.value)!="") {
        if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
          if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
        } else if (test!='R') { num = parseFloat(val);
          if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
          if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
            min=test.substring(8,p); max=test.substring(p+1);
            if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
      } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
    } if (errors) alert('The following error(s) occurred:\n'+errors);
    document.MM_returnValue = (errors == '');
} }
</script>
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
          <th scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../contactus/contactus.asp" class="navword">联系我们</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="backsystemlogin.asp" class="navword">管理入口</a></th>
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
    <th class="currentptable" scope="col"><a  href="../index.asp" class="dellink" 
onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">首页</a>-><a  href="backsystemlogin.asp" class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">用户登录</a></th>
  </tr>
</table>
</div>
	<div id="logincondiv"><form action="<%=MM_LoginAction%>" method="POST" onSubmit="MM_validateForm('username','','R','password','','R');return document.MM_returnValue"><table width="670" border="0" cellspacing="0" cellpadding="0">
    <tr>
     <th height="100" scope="col">&nbsp;</th>
    </tr>
  <tr>
    <th align="right" width="260" scope="col">用户名：</th>
    <th scope="col" align="left"><input name="username" type="text" id="username" size="20"></th>
  </tr>
  <tr>
    <th  align="right" scope="col">密码：</th>
    <th scope="col" align="left"><input name="password" type="password" id="password" size="20"></th>
  </tr>
  <tr>
  <td colspan="3" height="5">
  </td>
  </tr>
  <tr>
    <th scope="col">&nbsp;</th>
    <td align="left" ><input name="提交" type="submit" value="登录"><input name="重置" type="reset" id="重置"></td>
    
  </tr>
</table>
</form>
    
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
