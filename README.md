# salary_email
网易企业邮箱自动发送加密工资条邮件

脚本环境：
	
* python3
* selenium
* chrome 以及 chromedriver

实现逻辑：

* 利用selenium模拟登录至企业邮箱
* 跳转至写信页面
* 模拟选择编辑器内容为源码形式，并在更多选项中选择加密
* 从excel中读取邮件数据,循环发送
	![Selection_071.png](https://i.loli.net/2017/08/29/59a528d8c8638.png)




