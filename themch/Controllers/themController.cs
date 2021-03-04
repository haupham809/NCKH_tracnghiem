using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
using themch.Models;
namespace themch.Controllers
{
    public class themController : Controller
    {
        
        [HttpGet]
        public ActionResult themcauhoi()
        {
            return View();
        }
        [HttpPost]
        public PartialViewResult docfileword(HttpPostedFileBase file)
        {
            Application word = new Application();
            object miss = System.Reflection.Missing.Value;
            object path = Server.MapPath("~/Content/" + file.FileName);
            if (System.IO.File.Exists(path.ToString()))
            {
                System.IO.File.Delete(path.ToString());
            }

            file.SaveAs(path.ToString());

            object readOnly = true;
                object missing = System.Type.Missing;
                Document doc = word.Documents.OpenNoRepairDialog(ref  path,
                        ref miss, ref miss, ref miss, ref miss,
                        ref miss, ref miss, ref miss, ref miss,
                        ref miss, ref miss, ref miss, ref miss,
                        ref miss, ref miss, ref miss);

            string totalText = "";
                for (int i = 0; i < doc.Paragraphs.Count; i++)
                {
                    totalText += "\r\n" + doc.Paragraphs[i + 1].Range.Text.ToString();
                }

            object saveChanges = WdSaveOptions.wdPromptToSaveChanges;
            word.Documents.Close(saveChanges, missing, missing);
            List<Models.DapAn> dapan2 = new List<DapAn>();
                List<Models.CauHoi> cauhoi = new List<CauHoi>();
                for (int i = 0; i < totalText.Length; i++)
                {
                    if (totalText[i] == '$' && totalText[i + 1] == 'c' && totalText[i + 2] == '$')
                    {
                        int slcau = 0;
                        Models.CauHoi ch = new Models.CauHoi();
                        int sldapa = 0;
                        int slda = 0;
                        List<Models.DapAn> dapan = new List<DapAn>();

                        ch.CauHois1 = new List<DapAn>();
                        for (int j = i; j < totalText.Length; j++)
                        {

                            if ((totalText[j] == '$' && totalText[j + 1] == '*' && totalText[j + 2] == '$') || (totalText[j] == '$' && totalText[j + 1] == '$'))
                            {
                                slcau++;
                                Models.DapAn da = new DapAn();
                                if (slcau == 1)
                                {

                                    ch.NoiDubg1 = totalText.Substring(i + 3, j - i - 3);
                                    ch.HinhAnh1 = "";
                                    for (int z = 0; z < ch.NoiDubg1.Length - 2; z++)
                                    {
                                        if (ch.NoiDubg1[z] == '$' && ch.NoiDubg1[z + 1] == 'h' && ch.NoiDubg1[z + 2] == '$')
                                        {

                                            ch.HinhAnh1 = ch.NoiDubg1.Substring(z + 3, ch.NoiDubg1.Length - z - 3);
                                            ch.NoiDubg1 = ch.NoiDubg1.Substring(0, z);
                                        }

                                    }

                                }



                                for (int k = j + 2; k < totalText.Length; k++)
                                {


                                    if (totalText[j] == '$' && totalText[j + 1] == '*' && totalText[j + 2] == '$')
                                    {

                                        if (totalText[k] == '$' && totalText[k + 1] == '$')
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 3, k - j - 3);
                                            da.TrangThai1 = true;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            j = k - 1;
                                            ch.CauHois1.Add(da);


                                        }
                                        else if (totalText[k] == '$' && totalText[k + 1] == 'c' && totalText[k + 2] == '$')
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 3, k - 3 - j);
                                            da.TrangThai1 = true;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            sldapa++;
                                            j = k - 1;
                                            ch.CauHois1.Add(da);
                                            break;
                                        }
                                        else if (k == totalText.Length - 1)
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 3, totalText.Length - j - 3);
                                            da.TrangThai1 = true;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            sldapa++;
                                            j = totalText.Length - 1;
                                            ch.CauHois1.Add(da);
                                            break;
                                        }
                                    }

                                    else if (totalText[j] == '$' && totalText[j + 1] == '$')
                                    {

                                        if (totalText[k] == '$' && totalText[k + 1] == '$')
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 2, k - j - 2);
                                            da.TrangThai1 = false;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            j = k - 1;
                                            ch.CauHois1.Add(da);


                                        }
                                        else if (totalText[k] == '$' && totalText[k + 1] == '*' && totalText[k + 2] == '$')
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 2, k - j - 3);
                                            da.TrangThai1 = false;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            j = k - 1;
                                            ch.CauHois1.Add(da);

                                        }
                                        else if (totalText[k] == '$' && totalText[k + 1] == 'c' && totalText[k + 2] == '$')
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 2, k - j - 3);
                                            da.TrangThai1 = false;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            sldapa++;
                                            j = k - 1;
                                            ch.CauHois1.Add(da);
                                            break;
                                        }
                                        else if (k == totalText.Length - 1)
                                        {
                                            da.HinhAnh1 = "";
                                            da.NoiDung1 = totalText.Substring(j + 2, totalText.Length - j - 2);
                                            da.TrangThai1 = false;
                                            for (int z = 0; z < da.NoiDung1.Length - 2; z++)
                                            {
                                                if (da.NoiDung1[z] == '$' && da.NoiDung1[z + 1] == 'h' && da.NoiDung1[z + 2] == '$')
                                                {

                                                    da.HinhAnh1 = da.NoiDung1.Substring(z + 3, da.NoiDung1.Length - z - 3);
                                                    da.NoiDung1 = da.NoiDung1.Substring(0, z);
                                                }

                                            }
                                            sldapa++;
                                            ch.CauHois1.Add(da);
                                            j = totalText.Length - 1;
                                            break;
                                        }
                                    }

                                }



                            }



                            if (sldapa != 0)
                            {
                                cauhoi.Add(ch);
                                break;
                            }


                        }
                    }


                }
             
            
            
            
            return PartialView(cauhoi);

            }

    }
}