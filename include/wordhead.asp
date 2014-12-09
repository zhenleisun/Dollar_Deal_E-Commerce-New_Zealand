
<link href="images/nbase1.css" rel="stylesheet" type="text/css" />

<div class="top">
	<div class="top_t">
		<table width="961" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="337" height="94" align="left"><a href="index.asp"><img src="<%=weblogo%>" width="337" height="94"/></a></td>
			<td id="everyday-timer" width="450" background="images/ds_03.jpg" height="94" >
                <script>
                function TimeToHMS(seconds)
                {
                   sec = seconds % 60;
                   temp = ( seconds - sec ) / 60;
   
                   minute = temp % 60;
                   hour = (temp - minute) / 60;
   
                   if(!(isFinite(sec) && isFinite(minute) && isFinite(hour))) /* invalid time */
                   {
                      return ("");
                   }
   
                   time_str = hour;
                   time_str += " : ";   
                   time_str+=(minute<10)?("0"+minute):minute;
                   time_str+=" : ";
                   time_str+=(sec<10)?("0"+sec):sec;
   
                   return (time_str); 
                }
                function updateTimer1()
                {   
                    var currentHour = new Date().getHours();
                    if(currentHour>=10)
                        document.getElementById("everyday-timer").innerHTML=TimeToHMS(34*60*60-new Date().getHours()*60*60-new Date().getMinutes()*60-new Date().getSeconds());
                    else
                        document.getElementById("everyday-timer").innerHTML=TimeToHMS(10*60*60-currentHour*60*60-new Date().getMinutes()*60-new Date().getSeconds());

                   setTimeout('updateTimer1()', 1000);
                }

                updateTimer1();
                </script>
			</td>
			<td class="register" width="250" align="left" valign="top" background="images/ds_04.jpg">
            <table border="0" cellspacing="0" cellpadding="0">
              
              <tr>
                <td class="register-left">&nbsp;</td>
                <td class="register-right"><a class="top_kuai" href="buy.asp?action=show">Cart</a>&nbsp;&nbsp;&nbsp;<a class="top_kuai" href="user.asp">Login</a>&nbsp;/&nbsp;<a class="top_kuai" href="register.asp">Register</a></td>
              </tr>
            </table>            
            <table border="0" cellspacing="0" cellpadding="0">           
             <FORM action=research.asp method=post name=form2>
              <tr>
              <td  class="register-left"><INPUT class="wenbenkuang" name="searchkey" size=12/></td>
              
              <td  class="register-right"><INPUT name="Submit" style="background:url(images/ioc_anniu.gif) no-repeat; cursor:pointer; color:#FFF; padding:0 2px; border:none;" type="submit" value="Query"></td></tr>
              <TR style="display: none"> 
                     <TD><select name="anclassid">
                     <option value="0">classify</option> 
                      </select></TD>
					</TR>
              </FORM>
            </table>
            </td>
		  </tr>
	  </table>
	</div>
	<div class="top_c">
		<table width="961" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="94" height="32" align="center"><a href="index.asp" class="nav">Just In</a></td>
			<td width="76" align="center"><a href="class.asp?lx=big&anid=62" class="nav">Mens</a></td>
			<td width="106" align="center"><a href="class.asp?lx=big&anid=63" class="nav">Womens</a></td>
		    <td width="73" align="center"><a href="class.asp?lx=big&anid=64" class="nav">Boys</a></td>
		    <td width="84" align="center"><a href="class.asp?lx=big&anid=65" class="nav">Girls</a></td>
		    <td width="100" align="center"><a href="class.asp?lx=big&anid=66" class="nav">Footwear</a></td>
		    <td width="120" align="center"><a href="class.asp?lx=big&anid=67" class="nav">Accessories</a></td>
		    <td width="104" align="center"><a href="class.asp?lx=tejia" class="nav">Specials</a></td>
		    <td width="99" align="center"><a href="class.asp?lx=tuijian" class="nav">Recommend</a></td>
		    <td width="105" align="center"><a href="class.asp?lx=big&anid=68" class="nav">Gift Cards</a></td>
		  </tr>
	  </table>
	</div>
</div>