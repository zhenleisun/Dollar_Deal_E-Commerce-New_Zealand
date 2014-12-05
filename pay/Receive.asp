
<!--  
 * ====================================================================
 *
 *                 Receive.asp 由网银在线技术支持提供
 *
 *     本页面为支付完成后获取返回的参数及处理......
 *
 *
 * ====================================================================
-->

<!--#include file="MD5.asp"-->
<%
     ' 提取表单参数
	 ' 
	 ' v_oid
	 '      商户发送的v_oid定单编号     
	 ' v_pmode
	 '      支付方式（字符串）     
	 ' v_pstatus 
	 '      支付状态
     '      20（支付成功）
     '      30（支付失败）
	 ' v_pstring
	 '      支付结果信息 
     '      支付完成（当v_pstatus=20时）；
     '      失败原因（当v_pstatus=30时）；
     ' v_md5str
     '      Md5校验串
	 ' v_amount
	 '      订单实际支付金额
	 ' v_moneytype 
	 '      订单实际支付币种
	 ' remark1 
	 '      备注字段1
	 ' remark2 
	 '      备注字段2
	 ' key
	 '      私钥值，商户可上chinabank后台自行设定
	 '
	 '/
v_oid=request("v_oid")
v_pmode=request("v_pmode")
v_pstatus=request("v_pstatus")
v_pstring=request("v_pstring")
v_amount=request("v_amount")
v_moneytype=request("v_moneytype")
remark1=request("remark1")
remark2=request("remark2")
v_md5str=request("v_md5str")

key="32432423423423423"
  ' 请把key改为你的商户私匙

if request("v_md5str")="" then
	response.Write("v_md5str：空值")
	response.end
end if


'md5校验

text = v_oid&v_pstatus&v_amount&v_moneytype&key
md5text = Ucase(trim(md5(text)))


'按md5检验情况输出结果 Ucase转换为大写
if md5text<>v_md5str then
  response.write("MD5 error")
else
  '逻辑处理
  if v_pstatus=20 then
	'支付成功
  else
	'支付失败
  end if

'提示：仅是对校验码校验通过不表示该支付结果是成功只意味着该信息是由网银传回
'校验成功需对传回的v_pstatus参数做判断，其中20都意味着支付成功，30表示支付失败
'如果商户涉及实时售卡，请对返回的金额与数据库中原始金额做大小判断，以防恶意行为

   
'-----------------------------------------------
end if
%>






<!--
以下是打印出所有接收数据的结果，供编程人员参考
-->
<table width="93%" border="0">
  <tr> 
    <td> <p><b><font color="#FF0000">提示：</font> 您网上在线支付情况反馈如下：</b><br>
        此次交易编号： <%=v_oid%></p>
      <p> 
        <%if v_pstatus=20 then
								zhuangtai = "在线支付已经支付成功"
								%>
        在线支付已经支付成功 
        <%elseif v_pstatus=30 then
								zhuangtai = "在线支付失败!"
								%>
        在线支付失败! 
        <%end if%>
      </p>

        在线支付结果：<%=v_pstring%> </p>
      <p> 您所使用的卡为：<%=v_pmode%></p>
<strong>  请及时与我们联系提货，我们也会在第一时间内与你联系!</strong>



      </p></td>
  </tr>
</table>
