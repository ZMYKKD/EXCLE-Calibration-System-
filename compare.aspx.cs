using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using wapoffice.App_Code;
using System.Text.RegularExpressions;
using System.Reflection;


public partial class newland_tuanxian_compare : System.Web.UI.Page
{
    public DataTable Mydt { get; set; }
    public DataTable Fandt { get; set; }
    public bool mainKg { get; set; }
    public string message { get; set; }
    protected void Page_Load(object sender, EventArgs e)
    {
//        tools.setIP("学平险校验");
        var v_operno = "";
        if (string.IsNullOrEmpty(Request["opid"]))
        {
            if (Session["operTime"] == null)
            {
                Response.Write("<Script language=JavaScript>alert('请从学平险数据校验系统欢迎页登录');window.location.href='http://10.187.23.2:8081/newland/tuanxian/compare_list.aspx';</Script>");
                Response.End();
            }
            else
            {
                v_operno = Session["operTime"].ToString().Trim();
            }
        }
        else
        {
            Session["operTime"] = Request["opid"].Trim();
            v_operno = Request["opid"].Trim();
        }

        var sqlstr = "select * from chq_tx_xpxfan order by fan";
        Fandt = DbHelperInfor76.DbHelperInfor.ExecuteDataTable(sqlstr);

        if (Request.IsPostBack())
        {
            mainKg = true;  //总控制
            #region 导入
            if (Request["isok"] == "开始")
            {
                HttpPostedFile f = this.Request.Files[0];
                string fname = f.FileName;

                /* startIndex */
                int index = fname.LastIndexOf(".") + 1;
                /* length */
                int len = fname.Length - index;

                fname = DateTime.Now.ToFileTime().ToString() + "." + fname.Substring(index, len);
                fname = this.Server.MapPath("/newland/tuanxian/fileup/" + fname);
                /* save to server */
                //Response.Write(this.Server.MapPath("/cqtb/kp/2015mzkp/file/" + fname));
                f.SaveAs(fname);
               var v_oldname = ExcelHelper.GetSheetNames(fname);
               if (v_oldname.IndexOf("Sheet1") < 0)
               {
                   //删除文件
                   FuncDelfuile(fname);
                   Response.Write("<script language='javascript'>	alert('Excel文件标签1名称不为Sheet1，请核实，或重新下载标准模板！');history.back()	  </script>");
                   Response.End();
               }


                

                DataTable dt = new DataTable();
                dt = ExcelHelper.InputFromExcel(fname, "Sheet1");
                 
                var v_agex = 0;
                var v_agey = 99;
                var v_fan = Request["sele_fan"].Trim();
                if (v_fan != "0")
                {
                    sqlstr = "select agex,agey from chq_tx_xpxfan where fan='" + v_fan + "'";
                    var agedt =  DbHelperInfor76.DbHelperInfor.ExecuteDataTable(sqlstr);
                    if (agedt != null)
                    {
                        v_agex = Convert.ToInt32(agedt.Rows[0]["agex"].ToString());
                        v_agey = Convert.ToInt32(agedt.Rows[0]["agey"].ToString());
                    }
                }
                
                //Response.Write(DateTime.ParseExact("20150306","yyyyMMdd",null));
                xpx xpxM = new xpx();
                Ctxt ctxt = new Ctxt();
                Mydt = new DataTable();
                Mydt.Columns.Add("apname", typeof(string));
                Mydt.Columns.Add("pname", typeof(string));
                Mydt.Columns.Add("pid", typeof(string));
                Mydt.Columns.Add("school", typeof(string));
                Mydt.Columns.Add("bclass", typeof(string));
                Mydt.Columns.Add("bthdate", typeof(string));
                Mydt.Columns.Add("payseq", typeof(string));
                Mydt.Columns.Add("begdate", typeof(string));
                Mydt.Columns.Add("enddate", typeof(string));
                Mydt.Columns.Add("prelname", typeof(string));
                Mydt.Columns.Add("sex", typeof(string));
                Mydt.Columns.Add("tbtype", typeof(string));
                Mydt.Columns.Add("telno", typeof(string));
                Mydt.Columns.Add("apid", typeof(string));
                Mydt.Columns.Add("bz", typeof(string));
                int x0 = 0;
                int x1 = 0;
                int x2 = 0;
                int x3 = 0;
                int x4 = 0;
                int x5 = 0;
                int x6 = 0;
                int x7 = 0;
                int x8 = 0;
                int x9 = 0;
                int x10 = 0;
                int x11 = 0;
                int x12 = 0;
                int x13 = 0;
                int xfan = 0;
                int xid = 0;
                int xadd = 0;
                int xaddd = 0;
                int xadddd = 0;
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    xpxM.bz = "";
                    //投保人
                    ctxt = Ifempty(dt.Rows[i][0]);
                    if (ctxt.txt.ToString().Trim().Replace(" ", "").IndexOf("家长") >= 0)
                    {
                        xpxM.bz += "投保人姓名出错<br>" ;
                        x0++;
                    }
                    xpxM.apname = ctxt.txt;
                    x0 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "投保人为空<br>" : "";
                    //被保人
                    ctxt = Ifempty(dt.Rows[i][1]);
                    xpxM.pname = ctxt.txt;
                    x1 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "被保人为空<br>" : "";
                    //被保人身份证
                    ctxt = Ifid(dt.Rows[i][2]);
                    //if (string.IsNullOrEmpty(ctxt.txt) || ctxt.txt == "")
                    //{
                    //    x2 ++;
                    //    xpxM.bz += "被保人身份证为空<br>";
                    //}
                    xpxM.pid = ctxt.txt;
                    x2 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "被保人身份证有误<br>" : "";
                    if (dt.Rows[i][2]!=null && dt.Rows[i][1] !=null) 
                    {
                        N2n n2n = new N2n();
                        n2n = Ifname2name(dt.Rows[i][2], dt.Rows[i][1]);
                        if (n2n.isok)
                        {
                            xpxM.bz += "被保人姓名:"+n2n.sysname+"与导入不符<br>";
                            x2 += 1;
                        }
                    }
                    //学校
                    ctxt = Ifempty(dt.Rows[i][3]);
                    xpxM.school = ctxt.txt;
                    x3 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "学校为空<br>" : "";
                    //班级
                    ctxt = Ifempty(dt.Rows[i][4]);
                    xpxM.bclass = ctxt.txt;
                    x4 += ctxt.count;
                    //被保人生日
                    ctxt = IfdateY4md(dt.Rows[i][5]);
                    xpxM.bthdate = ctxt.txt;
                    x5 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "被保人生日有误<br>" : "";
                    if (!IfFan(dt.Rows[i][5], dt.Rows[i][7], v_agex, v_agey))
                    {
                        xpxM.bz += "被保人年龄不在投保范围<br>";
                        xfan++;
                    }

                    if (!AgeDaYu18(dt.Rows[i][13].ToString().Trim().Substring(6, 8), dt.Rows[i][7]))
                    {
                        xpxM.bz += "投保人年龄小于18岁<br>";
                        xadd++;
                    }
                    if ((dt.Rows[i][0].ToString().Trim() == dt.Rows[i][1].ToString().Trim()) &&(dt.Rows[i][2].ToString().Trim()!=dt.Rows[i][13].ToString().Trim()))
                    {
                        xpxM.bz += "投保或被保人姓名输入重复<br>";
                        xaddd++;
                    }
                    if ((dt.Rows[i][0].ToString().Trim() != dt.Rows[i][1].ToString().Trim()) && (dt.Rows[i][2].ToString().Trim() == dt.Rows[i][13].ToString().Trim()))
                    {
                        xpxM.bz += "投保或被保人身份证输入重复<br>";
                        xadddd++;
                    }


                    if (dt.Rows[i][2].ToString() != "" || !string.IsNullOrEmpty(dt.Rows[i][2].ToString()))
                    {
                        if (dt.Rows[i][2].ToString().Trim().Substring(6, 8) != dt.Rows[i][5].ToString().Trim())
                        {
                            xpxM.bz += "被保人出生日期与身份证不同<br>";
                            xid++;
                        }
                    }
                        
                    //受益人
                    ctxt = Ifempty(dt.Rows[i][6]);
                    xpxM.payseq = ctxt.txt;
                    x6 += ctxt.count;
                    //保险起期
                    ctxt = IfdateY4md(dt.Rows[i][7]);
                    xpxM.begdate = ctxt.txt;
                    x7 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "保险起期为空<br>" : "";

                    
                    //保险止期
                    ctxt = IfdateY4md(dt.Rows[i][8]);
                    xpxM.enddate = ctxt.txt;
                    x8 += ctxt.count;
                    //投保人与被保人关系
                    ctxt = Ifempty(dt.Rows[i][9]);
                    xpxM.prelname = ctxt.txt;
                    x9 += ctxt.count;
                    //性别
                    ctxt = Ifempty(dt.Rows[i][10]);
                    xpxM.sex = ctxt.txt;
                    x10 += ctxt.count;
                    //投保类型
                    ctxt = Iftbtype(dt.Rows[i][11]);
                    xpxM.tbtype = ctxt.txt;
                    x11 += ctxt.count;
                    //投保人电话号码
                    ctxt = Iftelphone(dt.Rows[i][12]);
                    xpxM.telno = ctxt.txt;
                    x12 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "投保人电话号码有误<br>" : "";
                    //投保人身份证号码
                    ctxt = Ifid(dt.Rows[i][13]);
                    xpxM.apid = ctxt.txt;
                    x13 += ctxt.count;
                    xpxM.bz += ctxt.count > 0 ? "投保人身份证号码有误<br>" : "";
                    if (dt.Rows[i][13] != null && dt.Rows[i][0] != null)
                    {
                        N2n n2n = new N2n();
                        n2n = Ifname2name(dt.Rows[i][13], dt.Rows[i][0]);
                        if (n2n.isok)
                        {
                            xpxM.bz += "投保人姓名:" + n2n.sysname+ "与导入不符<br>";
                            x13 += 1;
                        }
                    }
                    DataRow row = Mydt.NewRow();
                        row["apname"] = xpxM.apname;
                        row["pname"] = xpxM.pname;
                        row["pid"] = xpxM.pid;
                        row["school"] = xpxM.school;
                        row["bclass"] = xpxM.bclass;
                        row["bthdate"] = xpxM.bthdate;
                        row["payseq"] = xpxM.payseq;
                        row["begdate"] = xpxM.begdate;
                        row["enddate"] = xpxM.enddate;
                        row["prelname"] = xpxM.prelname;
                        row["sex"] = xpxM.sex;
                        row["tbtype"] = xpxM.tbtype;
                        row["telno"] = xpxM.telno;
                        row["apid"] = xpxM.apid;
                        row["bz"] = xpxM.bz;
                        Mydt.Rows.Add(row);

                }
                //删除文件
                FuncDelfuile(fname);
                if (x0 > 0)
                {
                    mainKg = false;
                    message += "投保人信息有误 ｜ ";
                }
                if (x1 > 0)
                {
                    mainKg = false;
                    message += "被保人有空 ｜ ";
                }
                if (x2 > 0)
                {
                    mainKg = false;
                    message += "被保人身份证号有误 ｜ ";
                }
                if (x3 > 0)
                {
                    mainKg = false;
                    message += "学校有空 ｜ ";
                }
                if (x5 > 0)
                {
                    mainKg = false;
                    message += "被保人生日有误或为空 ｜ ";
                }
                if(x7>0)
                {
                    mainKg = false;
                    message += "保险起期为空 ｜ ";
                }
                if (x12 > 0)
                {
                    mainKg = false;
                    message += "投保人电话有误或为空 ｜ ";
                }
                if (x13 > 0)
                {
                    mainKg = false;
                    message += "被保人身份证号有误 ｜ ";
                }
                if (xfan > 0)
                {
                    mainKg = false;
                    message += "被保人年龄不在投保范围 ｜ ";
                }
                if (xadd>0)
                {
                    mainKg = false;
                    message += "投保人年龄小于18岁 ｜ ";
                }
                if (xaddd > 0)
                {
                    mainKg = false;
                    message += "投保或被保人姓名输入重复 ｜ ";
                }
                if (xadddd > 0)
                {
                    mainKg = false;
                    message += "投保或被保人身份证输入重复 ｜ ";
                }
                if (xid > 0)
                {
                    mainKg = false;
                    message += "被保人出生日期与身份证不同 ｜ ";
                }
                if (mainKg)
                {
                    message += "通过";
                }

                //if (message != "通过")
                //{
                //    ExcelHelper.ExportExcelDT(Mydt, "参考错误提示模板");
                //}

            }


            #endregion
        }
    }
    //删除文件
    private void FuncDelfuile(string fname)
    {
        if (System.IO.File.Exists(fname))
        {
            try
            {
                System.IO.File.Delete(fname);
            }
            catch
            {
                Response.Write("<script language='javascript'>	alert('文件删除失败！');	  </script>");
            }
        }

    }
    //是否为空
    public Ctxt Ifempty(object str)
    {
        Ctxt ctxt = new Ctxt();

        if (str == null || string.IsNullOrEmpty(str.ToString()) || str.ToString().Length <1)
        {
            ctxt.txt = "";
            ctxt.count = 1;
        }
        else
        {
            ctxt.txt = str.ToString().Trim();
            ctxt.count = 0;
        }
        return ctxt;
    }
    //投保人电话号码
    public Ctxt Iftelphone(object str)
    {
        Ctxt ctxt = new Ctxt();

        if (str == null || string.IsNullOrEmpty(str.ToString()))
        {
            ctxt.txt = "";
            ctxt.count = 1;
        }
        else
        {
            if (IsValidPhoneAndMobile(str.ToString()))
            {
                ctxt.txt = str.ToString().Trim();
                ctxt.count = 0;
            }
            else
            {
                ctxt.txt = str.ToString().Trim();
                ctxt.count = 1;
            }
        }
        return ctxt;
    }
    //投保类型
    public Ctxt Iftbtype(object str)
    {
        Ctxt ctxt = new Ctxt();

        if (str == null || string.IsNullOrEmpty(str.ToString()))
        {
            ctxt.txt = "1";
            ctxt.count = 1;
        }
        else
        {
            ctxt.txt = str.ToString().Trim();
            ctxt.count = 0;
        }
        return ctxt;
    }
    //性别
    public Ctxt Ifsex(object str)
    {
        Ctxt ctxt = new Ctxt();

        if (str == null || string.IsNullOrEmpty(str.ToString()))
        {
            ctxt.txt = "1";
            ctxt.count = 1;
        }
        else
        {
            if (str.ToString().Trim() == "男")
            {
                ctxt.txt = "1";
            }
            else if (str.ToString().Trim() == "女")
            {
                ctxt.txt = "2";
            }
            else
            {
                ctxt.txt = str.ToString().Trim();
            }
            ctxt.count = 0;
        }
        return ctxt;
    }

    //身份证判断
    public Ctxt Ifid(object str)
    {
        Ctxt ctxt = new Ctxt();

        if (str == null || string.IsNullOrEmpty(str.ToString()))
        {
            ctxt.txt = "";
            ctxt.count = 0;
        }
        else
        {
            if (CheckIDCard(str.ToString().Trim()))
            {
                ctxt.txt = str.ToString().Trim();
                ctxt.count = 0;
            }
            else
            {
                ctxt.txt = str.ToString().Trim();
                ctxt.count = 1;
            }
        }
        return ctxt;
    }
    //日期处理
    public Ctxt IfdateY4md(object str)
    {
        DateTime result;
        Ctxt ctxt = new Ctxt();
        if (str == null || string.IsNullOrEmpty(str.ToString()) )
        {
            ctxt.txt = str.ToString(); 
            ctxt.count = 1;
        }
        else

        {
            if ( str.ToString().Length != 8)
            {
                ctxt.txt = str.ToString();
                ctxt.count = 1;
            }
        else if (!DateTime.TryParse(str.ToString().Substring(0, 4) + "-" + str.ToString().Substring(4, 2) + "-" + str.ToString().Substring(6, 2) + " 00:00:00", out result))
                {
                    ctxt.txt = str.ToString();
                    ctxt.count = 1;
                }
         if (str.ToString().Length == 8 && tools.IsNum(str.ToString()))
        {
                ctxt.txt = str.ToString();
                ctxt.count = 0;
        }
        else
        {
                //if (tools.IsDate(str.ToString()))
                //{
                //    DateTime dt = Convert.ToDateTime(str.ToString());
                //   ctxt.txt = string.Format("{0:yyyyMMdd}", dt);
                //   ctxt.count = 0;
                //}
                //else
                //{
                ctxt.txt = str.ToString();
                ctxt.count = 1;
                //}
        }
        }
        return ctxt;
    }

    //投保人年龄>18?
    public bool AgeDaYu18(object age,object bdate)
    {
        int age1=0;
        bool cs;
        if (age == null || string.IsNullOrEmpty(age.ToString()))
        {
            cs = false;
        }
        else
        {
            if (age.ToString().Length == 8 && tools.IsNum(age.ToString()))
            {
                DateTime result;//将身份证时间转化为时间格式
                if (!DateTime.TryParse(age.ToString().Substring(0, 4) + "-" + age.ToString().Substring(4, 2) + "-" + age.ToString().Substring(6, 2) + " 00:00:00", out result))
                    cs = false;
                else
                {
                    DateTime birthdate = DateTime.ParseExact(age.ToString().Substring(0, 4) + "-" + age.ToString().Substring(4, 2) + "-" + age.ToString().Substring(6, 2) + " 00:00:00", "yyyy-MM-dd 00:00:00", null);
                    DateTime now = DateTime.ParseExact(bdate.ToString().Substring(0, 4) + "-" + bdate.ToString().Substring(4, 2) + "-" + bdate.ToString().Substring(6, 2) + " 00:00:00", "yyyy-MM-dd 00:00:00", null);
                    age1 = now.Year - birthdate.Year;
                    if (now.Month < birthdate.Month || (now.Month == birthdate.Month && now.Day < birthdate.Day))
                    {
                        age1--;
                    }
                    age1 = age1 < 0 ? 0 : age1;
                    if (age1 >= 0 && age1 <= 18)
                    {
                        cs = false;
                    }
                    else
                    {
                        cs = true;
                    }
                }
            }
            else
            {
                cs = false;
            }
        }

        return cs;
    }




    //方案投保年龄判断
    public bool IfFan(object str,object bdate,int v_agex,int v_agey)
    {
        int age = 0;
        bool v_fan;
        if (str == null || string.IsNullOrEmpty(str.ToString()))
        {
            v_fan = false;
        }
        else
        {
            if (str.ToString().Length == 8 && tools.IsNum(str.ToString()) && bdate.ToString().Length == 8 && tools.IsNum(bdate.ToString()))
            {
                //判断月
                //if (str.ToString().Substring(4, 2).ToInt32() < 1 || str.ToString().Substring(4, 2).ToInt32() > 12)
                //{
                //    v_fan = false;
                //}
                //else if (str.ToString().Substring(6, 2).ToInt32() < 1 || str.ToString().Substring(6, 2).ToInt32() > 31)
                //{
                //    v_fan = false;
                //}
                DateTime result; 
                if (!DateTime.TryParse(str.ToString().Substring(0,4)+"-"+str.ToString().Substring(4,2)+"-"+str.ToString().Substring(6,2)+" 00:00:00", out result))
                    v_fan = false;
                else
                {

                    //DateTime birthdate = Convert.ToDateTime(str.ToString());
                    //DateTime birthdate = DateTime.ParseExact(str.ToString(), "yyyy-MM-dd 00:00:00", null);
                    //DateTime now = DateTime.ParseExact(bdate.ToString(), "yyyy-MM-dd 00:00:00", null);
                    DateTime birthdate = DateTime.ParseExact(str.ToString().Substring(0, 4) + "-" + str.ToString().Substring(4, 2) + "-" + str.ToString().Substring(6, 2) + " 00:00:00", "yyyy-MM-dd 00:00:00", null);
                    DateTime now = DateTime.ParseExact(bdate.ToString().Substring(0, 4) + "-" + bdate.ToString().Substring(4, 2) + "-" + bdate.ToString().Substring(6, 2) + " 00:00:00", "yyyy-MM-dd 00:00:00", null);
                    age = now.Year - birthdate.Year;
                    if (now.Month < birthdate.Month || (now.Month == birthdate.Month && now.Day < birthdate.Day))
                    {
                        age--;
                    }
                    age = age < 0 ? 0 : age;
                    if (age >= v_agex && age <= v_agey)
                    {
                        v_fan = true;
                    }
                    else
                    {
                        v_fan = false;
                    }
                }
            }
            else
            {
                v_fan = false;
            }
        }
        return v_fan;
    }
    //文件名字和系统名字对比
    public N2n Ifname2name(object strid, object strname)
    {
        N2n n2n = new N2n();
        var tmp = false;

        var strsql = "select name from test3_person where id_id15='" + strid.ToString().Trim() + "' and name  is not null";
        var sysname = DbHelperInfor76.DbHelperInfor.ExecuteSqlScalar(strsql);
        if (sysname != null)
        {
            if (sysname.ToString().Trim().Replace(" ", "") != strname.ToString().Trim())
            {
                tmp = true;
                n2n.sysname = sysname.ToString().Trim().Replace(" ", "");
            }
        }
        n2n.isok = tmp;
        
        return n2n;
    }


    public class N2n
    {
        public bool isok { get; set; }
        public string sysname { get; set; }
    }
    public class Ctxt
    {
        public string txt { get; set; }
        public int count { get; set; }
    }
    public class xpx
    {
        public string apname { get; set; }
        public string pname { get; set; }
        public string pid { get; set; }
        public string school { get; set; }
        public string bclass { get; set; }
        public string bthdate { get; set; }
        public string payseq { get; set; }
        public string begdate { get; set; }
        public string enddate { get; set; }
        public string prelname { get; set; }
        public string sex { get; set; }
        public string tbtype { get; set; }
        public string telno { get; set; }
        public string apid { get; set; }
        public string bz { get; set; }
    }

    /// <summary>
    /// 电话有效性（固话和手机 ）
    /// </summary>
    /// <param name="strVla"></param>
    /// <returns></returns>
    public static bool IsValidPhoneAndMobile(string number)
    {
        Regex rx = new Regex(@"^(\(\d{3,4}\)|\d{3,4}-)?\d{7,8}$|^(13|15|17|18|19|14|16|98|92)\d{9}$", RegexOptions.None);
        Match m = rx.Match(number);
        return m.Success;
    }
    /// 身份证验证
    /// </summary>
    /// <param name="Id">身份证号</param>
    /// <returns></returns>
    public bool CheckIDCard(string Id)
    {
        if (Id.Length == 18)
        {
            bool check = CheckIDCard18(Id);
            return check;
        }
        else if (Id.Length == 15)
        {
            bool check = CheckIDCard15(Id);
            return check;
        }
        else
        {
            return false;
        }
    }
    /// <summary>
    /// 18位身份证验证
    /// </summary>
    /// <param name="Id">身份证号</param>
    /// <returns></returns>
    private bool CheckIDCard18(string Id)
    {
        long n = 0;
        if (long.TryParse(Id.Remove(17), out n) == false || n < Math.Pow(10, 16) || long.TryParse(Id.Replace('x', '0').Replace('X', '0'), out n) == false)
        {
            return false;//数字验证
        }
        string address = "11x22x35x44x53x12x23x36x45x54x13x31x37x46x61x14x32x41x50x62x15x33x42x51x63x21x34x43x52x64x65x71x81x82x91";
        if (address.IndexOf(Id.Remove(2)) == -1)
        {
            return false;//省份验证
        }
        if (Regex.IsMatch(Id.Substring(17), "[a-z]"))
        {
            return false;//小写
        }
        string birth = Id.Substring(6, 8).Insert(6, "-").Insert(4, "-");
        DateTime time = new DateTime();
        if (DateTime.TryParse(birth, out time) == false)
        {
            return false;//生日验证
        }
        string[] arrVarifyCode = ("1,0,x,9,8,7,6,5,4,3,2").Split(',');
        string[] Wi = ("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2").Split(',');
        char[] Ai = Id.Remove(17).ToCharArray();
        int sum = 0;
        for (int i = 0; i < 17; i++)
        {
            sum += int.Parse(Wi[i]) * int.Parse(Ai[i].ToString());
        }
        int y = -1;
        Math.DivRem(sum, 11, out y);
        if (arrVarifyCode[y] != Id.Substring(17, 1).ToLower())
        {
            return false;//校验码验证
        }
        return true;//符合GB11643-1999标准
    }
    /// <summary>
    /// 15位身份证验证
    /// </summary>
    /// <param name="Id">身份证号</param>
    /// <returns></returns>
    private bool CheckIDCard15(string Id)
    {
        long n = 0;
        if (long.TryParse(Id, out n) == false || n < Math.Pow(10, 14))
        {
            return false;//数字验证
        }
        string address = "11x22x35x44x53x12x23x36x45x54x13x31x37x46x61x14x32x41x50x62x15x33x42x51x63x21x34x43x52x64x65x71x81x82x91";
        if (address.IndexOf(Id.Remove(2)) == -1)
        {
            return false;//省份验证
        }
        string birth = Id.Substring(6, 6).Insert(4, "-").Insert(2, "-");
        DateTime time = new DateTime();
        if (DateTime.TryParse(birth, out time) == false)
        {
            return false;//生日验证
        }
        return true;//符合15位身份证标准
    }

}

