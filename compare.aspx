<%@ Page Language="C#" AutoEventWireup="true" CodeFile="compare.aspx.cs" Inherits="newland_tuanxian_compare" %>


<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- 上述3个meta标签*必须*放在最前面，任何其他内容都*必须*跟随其后！ -->
    <title>学平险数据校验1111</title>
     <!-- Bootstrap -->
    <link href="/bootstrap/css//bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://cdn.bootcss.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
       <script>
           function a() {

               var fname = document.getElementById("excelfile").value;
               if (fname != "") {
                   return true;
               }
               else {
                   alert("未选择任何文件！");
                   return false;
               }
           }
       </script>

       <script>
           function exportXls() {
               window.clipboardData.setData("Text", document.all('tab1').outerHTML);
               try {
                   var ExApp = new ActiveXObject("Excel.Application")
                   var ExWBk = ExApp.workbooks.add()
                   var ExWSh = ExWBk.worksheets(1)
                   ExApp.DisplayAlerts = false
                   ExApp.visible = true
               } catch (e) {
                   alert("您的电脑没有安装Microsoft Excel软件！")
                   return false
               }
               ExWBk.worksheets(1).Paste;
           }
       </script>
</head>
<body>
<div  class="container">
<h3 class="page-header">学平险数据校验:临时通道111111111<small>v.1.9</small></h3>

    <form method="post" action="compare.aspx" enctype="multipart/form-data" class="form-inline">
    <div class="well">
        <div class="form-group">
        <label>上传文件</label>
        <input type="file" class="form-control"   name="excelfile" id="excelfile" />
         </div>

         <div class="form-group">
        <label>学平险方案：</label>
        <select name="sele_fan" class="form-control">
        <option value="0">单纯校验格式</option>
        <%
            for (int i = 0; i < Fandt.Rows.Count; i++)
            {
                %>
        <option value="<%=Fandt.Rows[i]["fan"] %>"><%=Fandt.Rows[i]["fan"] %> - <%=Fandt.Rows[i]["agex"] %>:<%=Fandt.Rows[i]["agey"] %></option>        
                <%
            }
             %>
        </select>
         

         <div class="form-group">
        <input type="submit" name="isok"  class="btn btn-success" value="开始" onclick="return a()"/>
        <button type="button" onclick="exportXls()" class="btn btn-primary">复制内容</button>
        <a href="compare_list.aspx" class="btn btn-default">返回</a>
        <%--<a href="http://10.81.128.102:8082/newland/tuanxian/compare.aspx" class="btn btn-danger" target="_blank">校验系统临时通道－分流</a>--%>
         </div>
  
    </div>
    <p><b>
    说明：Excel文件必须使用学平险专用模板 <a href="fileup/xpx_moban.xls">学平险专用模板下载</a><br />
    </b></p>
    <small> 系统处于测试期，有问题请即时联系信息技术部兰元庆: 63795500-5058 ｜ <b class="text-danger">复制内容：只能使用IE浏览器，<i>数据只能做参考</i>使用</b></small>
    </div>
  
    </form>
    <%
        if (Mydt!=null)
        {
         %>
         <label <%=mainKg ? "class='label-success'":"class='label-danger'" %>  ><%=message%></label>
                            <div style="overflow:scroll; height:600px">
    <table id="tab1" class="table table-bordered table-condensed" style="min-width:1500px;">
    <tr>
    <td>提示</td>
    <td>投保人</td>
<td>被保人</td>
<td>被保人证件类型</td>
<td>被保人ID</td>
<td>学校</td>
<td>班级</td>
<td>被保人生日</td>
<td>受益人</td>
<td>保险起期</td>
<td>保险止期</td>
<td>投被保人关系</td>
<td>性别</td>
<td>投保类型</td>
<td>电话</td>
        <td>投保人证件类型</td>
<td>投保人ID</td>
        <td>投保人港澳台通行证</td>
        <td>被保人港澳台通行证</td>

    </tr>
    <%for (int i = 0; i < Mydt.Rows.Count; i++)
      {
          if (Mydt.Rows[i][3] != "" || Mydt.Rows[i][5] != "" || Mydt.Rows[i][1] !="") { 
          
          
          %>
    <tr <%=Mydt.Rows[i][14] =="" ?"":"class='label-warning'" %>>
    <td class="text-danger"><%=Mydt.Rows[i][14].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][0].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][1].ToString().Trim()%></td>
    <td>[<%=Mydt.Rows[i][2].ToString().Trim()%>]</td>
    <td><%=Mydt.Rows[i][3].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][4].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][5].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][6].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][7].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][8].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][9].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][10].ToString().Trim()%></td>
    <td><%=Mydt.Rows[i][11].ToString().Trim()%></td>
    <td>[<%=Mydt.Rows[i][12].ToString().Trim()%>]</td>
    <td>[<%=Mydt.Rows[i][13].ToString().Trim()%>]</td>
    </tr>      
          <%
          }
      }%>
    </table>
                    </div>
    <%} %>

    <small>
    验校规则：1、投被保人身份证格式；2、投被人姓名写系统姓名是否一致；3、学校不能为空；4、被保人生日在方案要求日期内；5、电话号码
    常见错误提示：<br />
    问题一：'Sheet1$' 不是一个有效名称。请确认它不包含无效的字符或标点，且名称不太长。<br />
    答：Excel模板文件的数据标签名称不是"Sheet1"，需要修改标签；<br />
    更新：<br />
    2018-9-25：投保人年龄按投保日期差计划 －－主城v.1.8<br />
    2018-8-21：增加判断生日的月日核验v.1.6
    2018-8-17：增加结果表内容下载功能，只能使用IE浏览器－－渠道部v.1.5<br />
    2018-8-14：增加投保人姓名判断不能包含家长、保险起期不能为空－－契约部v.1.4<br />
    2018-8-7：客户姓名不符合，提示系统中的姓名（最终以生效提示为准）－－主城v.1.3<br />
    </small>
    </div>
  
</body>
</html>
