
<!--  
 * ====================================================================
 *
 *                 Receive.asp ���������߼���֧���ṩ
 *
 *     ��ҳ��Ϊ֧����ɺ��ȡ���صĲ���������......
 *
 *
 * ====================================================================
-->

<!--#include file="MD5.asp"-->
<%
     ' ��ȡ��������
	 ' 
	 ' v_oid
	 '      �̻����͵�v_oid�������     
	 ' v_pmode
	 '      ֧����ʽ���ַ�����     
	 ' v_pstatus 
	 '      ֧��״̬
     '      20��֧���ɹ���
     '      30��֧��ʧ�ܣ�
	 ' v_pstring
	 '      ֧�������Ϣ 
     '      ֧����ɣ���v_pstatus=20ʱ����
     '      ʧ��ԭ�򣨵�v_pstatus=30ʱ����
     ' v_md5str
     '      Md5У�鴮
	 ' v_amount
	 '      ����ʵ��֧�����
	 ' v_moneytype 
	 '      ����ʵ��֧������
	 ' remark1 
	 '      ��ע�ֶ�1
	 ' remark2 
	 '      ��ע�ֶ�2
	 ' key
	 '      ˽Կֵ���̻�����chinabank��̨�����趨
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
  ' ���key��Ϊ����̻�˽��

if request("v_md5str")="" then
	response.Write("v_md5str����ֵ")
	response.end
end if


'md5У��

text = v_oid&v_pstatus&v_amount&v_moneytype&key
md5text = Ucase(trim(md5(text)))


'��md5������������� Ucaseת��Ϊ��д
if md5text<>v_md5str then
  response.write("MD5 error")
else
  '�߼�����
  if v_pstatus=20 then
	'֧���ɹ�
  else
	'֧��ʧ��
  end if

'��ʾ�����Ƕ�У����У��ͨ������ʾ��֧������ǳɹ�ֻ��ζ�Ÿ���Ϣ������������
'У��ɹ���Դ��ص�v_pstatus�������жϣ�����20����ζ��֧���ɹ���30��ʾ֧��ʧ��
'����̻��漰ʵʱ�ۿ�����Է��صĽ�������ݿ���ԭʼ�������С�жϣ��Է�������Ϊ

   
'-----------------------------------------------
end if
%>






<!--
�����Ǵ�ӡ�����н������ݵĽ�����������Ա�ο�
-->
<table width="93%" border="0">
  <tr> 
    <td> <p><b><font color="#FF0000">��ʾ��</font> ����������֧������������£�</b><br>
        �˴ν��ױ�ţ� <%=v_oid%></p>
      <p> 
        <%if v_pstatus=20 then
								zhuangtai = "����֧���Ѿ�֧���ɹ�"
								%>
        ����֧���Ѿ�֧���ɹ� 
        <%elseif v_pstatus=30 then
								zhuangtai = "����֧��ʧ��!"
								%>
        ����֧��ʧ��! 
        <%end if%>
      </p>

        ����֧�������<%=v_pstring%> </p>
      <p> ����ʹ�õĿ�Ϊ��<%=v_pmode%></p>
<strong>  �뼰ʱ��������ϵ���������Ҳ���ڵ�һʱ����������ϵ!</strong>



      </p></td>
  </tr>
</table>