# ProjectWeb
基于Dreamweaver开发环境，搭建了课题组网站，实现了包括网页样式设计（CSS）、网页动态显示（图片轮播等）、后台操作管理系统、数据库管理。
应用流程如下
1.首先下载Dreamweaver CS6；下载与破解：http://jingyan.baidu.com/article/fcb5aff7ac492aedab4a7167.html

2.搭建IIS服务环境以及建立Dreamweaver站点；http://www.jb51.net/article/29787.html；http://www.adobe.com/cn/devnet/dreamweaver/articles/setup_asp.html；其中搭建IIS服务器有几项注意点：a.物理路径建议不要更改,b.网站绑定建议默认，c.无需添加网站使用Default Web Site即可.d.如果出现问题先站点文件给everyone完全权限，然后C:\Windows\temp文件夹给USER完全权限。

3.首先先在C:\inetpub\wwwroot\ProjectWebs\indeximages文件夹下替代ground.jpg（前台网页的背景，图片比例900:204）backsystem.jpg（后台网页的背景 图片比例900:204）schoolplace.png（学校地图 图片比例 665:300）教授.jpg（学术带头人照片 图片比例100:120）
然后在C:\inetpub\wwwroot\ProjectWebs\images中替换s1.jpg-s4.jpg,sc1.jpg-sc4.jpg（主页轮显图片 图片比例663:291）
4.打开C:\inetpub\wwwroot\ProjectWebs\index.asp，定位到141行到176行，张三教授替换为Boss的名字，机械电子系系主任，智能机器人所所长替换成Boss的职务，xxxxx......xxxxxx处替换Boss的个人简历（注意换行）
定位到209到236行，其中http://www.baidu.com/替换成你需要链接过去的网站，百度替换为网站名
定位到239到252行，其中&copy;&nbsp;Copyright&nbsp;2007-2014&nbsp;百度百度百度课题组版权所有。&copy是版权信息的意思，&nbsp;是空格的意思。把时间换掉课题组名称换掉。北京市西长安街174号中南海新华门|邮编：100017，换掉。电话：010-11111111|邮箱：xiaobai@163.com|推荐兼容模式打开，换掉。
定位到92-93行，<title>课题组</title> 课题组替换为主页或者课题组名称，具体什么效果试一下就知道了。<meta name="keywords" content="课题组">
content里面的内容是搜索引擎检索的关键词，搜索课题组就可以搜到本网站。
5.接着打开所有asp后缀的文件挨个改最下方的版权信息内容，如果有学术带头人的信息则一并更改。至此大致框架已经形成接下来修改网站样式。
6.打开C:\inetpub\wwwroot\ProjectWeb\main.css   定位到20到25行
.navbar {								//前缀.表示一个类前缀#表示一个ID。这个类用来改变导航栏表格的大小
	background-color: #3388ff;			//改变导航栏背景颜色
	height: 50px; 						//设定高度
	width: 900px;						//设定长度
	margin-top: -20px;					//设定位置
	font-weight: bold;					//导航栏字体粗细 
	
}
定位到28到34
.navword {							//设定导航栏中的字体样式
	font-family: "微软雅黑";		//设定字体类型
	font-size: 18px;				//大小
	font-weight: bold;				//粗细
	color: #FFF;					//颜色
	text-decoration: none;			//无下划线
}

.leadertable{						//设定学术带头人这一个单元格的大小以及这几个字的样式
	height: 50px;
	width: 220px;
	font-family:'微软雅黑'; 
	font-size:16px; 
	font-weight:bold;
	background-color:#61a3ff;
}	
.leaderintroductionword{			//设定Boss简介的字体格式
	text-align: left;
	font-family: "宋体";
	font-size: 14px;
	font-weight: normal;
	
}
定位到121到125
#footerwords {						//设定版权信息字体的格式
	font-family: "宋体";
	font-size: 14px;
	font-weight: normal;
}
定位到134到139
.footsong{							//设定网页正文的字体格式，网站中大部分的字体格式都是用这个类来表示的
	font-family:"宋体";		
	font-size: 14px;
	font-weight:normal;
	text-decoration:none;
	color:#000;}

7.基本到这里网站已经建立的差不多了，接下来就是通过后台管理系统向数据库里面写数据。然后来说一下后台管理系统的使用注意事项，首先不支持中文文件上传；文件名与上传的文件名称需一致；部分修改页面中有下拉列表选项，其中下拉列表不会随着信息的变化而变化需要重新选择，否则会默认选择第一个；密码的更改可以打开ProjectWeb.mdb数据库更改其中的login表里的username和password来实现；初始账号密码均为123456789。
