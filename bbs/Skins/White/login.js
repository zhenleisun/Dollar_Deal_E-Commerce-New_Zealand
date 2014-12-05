function SWLogin()
{
oLogin = document.getElementById("Login");
oIEBackground = document.getElementsByTagName("body");
oCSS3Background = document.getElementById("LoginBackground");
if (oLogin.style.display == "block")
{
oLogin.style.display = "none";
oIEBackground[0].style.filter = "";
oCSS3Background.style.opacity = "";
}
else
{
oLogin.style.display = "block";
oIEBackground[0].style.filter = "alpha(opacity=6)";
oCSS3Background.style.opacity = "0.06";
}
}
