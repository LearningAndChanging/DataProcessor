<!DOCTYPE html>
<html>
<head>
 <meta charset="utf-8" />
 <link rel="stylesheet" href="./database/css/style.css">
 <title>新材料学院数据库系统登陆界面</title>
 <style type="text/css">
 	.t1{
		font-size:16px;
	}
	.t2{
		font-weight:bold;
		font-size:20px;
	}
</style>
 <script>
 function init1(){
	if(myform1.username1.value=="")
	{
		alert("请输入用户名");
		//将光标移动到文本框中
		myform1.username1.focus();
		return false;
	}
	if (myform1.userpwd1.value=="")
	{
		alert("请输入密码");
		myform1.userpwd1.focus();
		return false;
	}
	if (myform1.code1.value=="")
	{
		alert("请输入验证码");
		myform1.code1.focus();
		return false;
	}
}
 function init2(){
    if(myform2.username2.value=="")
    {
        alert("请输入用户名");
        //将光标移动到文本框中
        myform2.username2.focus();
        return false;
    }
    if (myform2.userpwd2.value=="")
    {
        alert("请输入密码");
        myform2.userpwd2.focus();
        return false;
    }
    if (myform2.confirm2.value=="")
    {
        alert("请再输入一次密码");
        myform2.confirm2.focus();
        return false;
    }
    if (myform2.code2.value=="")
    {
        alert("请输入验证码");
        myform2.code2.focus();
        return false;
    }
}
</script>
</head>

<body>

<div class="cotn_principal">
  <div class="cont_centrar">
    <div class="cont_login">
      <div class="cont_info_log_sign_up">
        <div class="col_md_login">
          <div class="cont_ba_opcitiy">
            <h2>LOGIN</h2>
            <p>如果已经拥有账号，请直接登录。</p>
            <button class="btn_login" onClick="cambiar_login()" id="btn1">LOGIN</button>
          </div>
        </div>
        <div class="col_md_sign_up">
          <div class="cont_ba_opcitiy">
            <h2>SIGN UP</h2>
            <p>如果未拥有账号，请先注册。</p>
            <button class="btn_sign_up" onClick="cambiar_sign_up()" id="btn2">SIGN UP</button>
          </div>
        </div>
      </div>
      <div class="cont_back_info">
        <div class="cont_img_back_grey"> <img src="./database/images/background.jpg" alt="" /> </div>
      </div>
      <div class="cont_forms" >
        <div class="cont_img_back_"> <img src="./database/images/background.jpg" alt="" /> </div>
        <form action="./database/logincheck.php" method="post" onsubmit="return init1();" name="myform1" >
        <div class="cont_form_login"> <a href="#" onClick="ocultar_login_sign_up()" ><i class="material-icons">&#xE5C4;</i></a>
          <h2>LOGIN</h2>
          <input type="text" name="username1" id="username1" placeholder="UserName" />
          <input type="password" name="userpwd1" id="userpwd1" placeholder="Password" /><br></br>
          <span class="t1">验证码:</span>
		  <span><input type="text" name="code1" class="code" id="code1" placeholder="请输入验证码"></span>
		  <span><img src="./database/pic1.php" onClick="this.src='./database/pic1.php?nocache='+Math.random()" style="cursor:pointer" id="img1"></span>
          <button class="btn_login" >LOGIN</button>
          <span><input type = "hidden" name = "hidden" value = "hidden" /></span>
        </div>
        </form>
        <form action="./database/regcheck.php" method="post" onsubmit="return init2();" name="myform2" >
        <div class="cont_form_sign_up"> <a href="#" onClick="ocultar_login_sign_up()"><i class="material-icons">&#xE5C4;</i></a>
          <h2>SIGN UP</h2>
          <div class="t1">请选择所在课题组：
            <select name="groups" id="groups">
            <option value="yinjiang">银浆课题组</option>
            <option value="sanyuan">三元材料组</option>
            <option value="">option3</option>
            <option value="">option4</option>
            </select></div>
          <input type="text" name="username2" id="username2" placeholder="UserName（仅限使用中文姓名注册）" />
          <input type="password" name="userpwd2" id="userpwd2" placeholder="Password（密码长度不小于6位）" />
          <input type="password" name="confirm2" id="confirm2" placeholder="Confirm Password" /><br></br>
          <span class="t1">验证码:</span>
		  <span><input type="text" name="code2" class="code" id="code2" placeholder="请输入验证码"></span>
		  <span><img src="./database/pic2.php" onClick="this.src='./database/pic2.php?nocache='+Math.random()" id="img2"></span>
          <button class="btn_sign_up" >SIGN UP</button>
          <span><input type = "hidden" name = "hidden" value = "hidden" /></span>
        </div>
        </form>
      </div>
    </div>
  </div>
  <div class="t2">建议使用Chrome浏览器进行浏览，其他浏览器上均未进行测试。</div>
  <div class="t2">如遇使用bug，请联系姜行进行解决。</div>
</div>

<script src="./database/js/index.js"></script>
<script type="text/javascript">

</script>
<div style="text-align:center;">
</div>


</body>
</html>
