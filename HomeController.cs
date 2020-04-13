using ImportData.Base;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;


namespace ImportData
{
    [ApiController]
    [Route("[controller]")]
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private AppDb _db;

        public HomeController(AppDb db, IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            _db = db;
        }

        /// <summary>
        /// 储值卡
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("pet-card")]
        public async Task<string> PetCard(IFormFile excelfile)
        {
            /*
             * TODO  
             * 1.读取表格数据
             * 2.查找名字
             * 3.插入数据
             */
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            DataTable dt = GetPetCardDT();
            var exceldt = GetExcelDT(excelfile, file, dt);

            List<DisposeDBModel> list = new List<DisposeDBModel>();
            string NotUserInfo = null;
            try
            {

                int skipCount = 0;
                using (_db)
                {
                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();

                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        string strName = exceldt.Rows[i]["姓名"].ToString();

                        cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where `Name`='" + strName + "'";
                        var result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        if (!result.Any())
                        {
                            cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where Nickname='" + strName + "'";
                            result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        }
                        if (result.Any())
                        {
                            list.Add(new DisposeDBModel
                            {
                                Balance = decimal.Parse(exceldt.Rows[i]["剩余金额"].ToString()),
                                Discriminator = exceldt.Rows[i]["会员卡"].ToString() == "专属卷" ? "GroupSaving" : "CommonSaving",
                                OwnerId = result[0].Id
                            });
                        }
                        else
                        {
                            NotUserInfo = NotUserInfo + "编号:" + exceldt.Rows[i]["编号"].ToString() + " 姓名:" + strName + " 会员卡类型:" + exceldt.Rows[i]["会员卡"].ToString() + " 剩余金额:" + exceldt.Rows[i]["剩余金额"].ToString() + "\n";
                            skipCount++;
                        }
                    }
                }
                int count = 0;
                using (_db)
                {

                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();
                    foreach (var item in list)
                    {
                        await InsertAsync(_db, item);
                        count++;
                    }
                }

                using (StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "储值卡.txt"))))
                {
                    txtFile.WriteLine(NotUserInfo);
                }
                return "导入:" + count + " 跳出:" + skipCount;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        /// <summary>
        /// 成人减脂
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("crjz-card")]
        public async Task<string> CrjzCard(IFormFile excelfile)
        {
            /*
             * TODO  
             * 1.读取表格数据
             * 2.查找名字
             * 3.插入数据
             */
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            DataTable dt = GetCrjzCardDT();
            var exceldt = GetExcelDT(excelfile, file, dt, 1);

            List<DisposeDBModel> list = new List<DisposeDBModel>();
            string NotUserInfo = null;
            try
            {

                int skipCount = 0;
                using (_db)
                {
                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();

                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        string strName = exceldt.Rows[i]["姓名"].ToString();

                        cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where `Name`='" + strName + "'";
                        var result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        if (!result.Any())
                        {
                            cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where Nickname='" + strName + "'";
                            result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        }
                        if (result.Any())
                        {
                            list.Add(new DisposeDBModel
                            {
                                Balance = decimal.Parse(exceldt.Rows[i]["剩余次数"].ToString()),
                                Discriminator = exceldt.Rows[i]["会员卡"].ToString(),
                                OwnerId = result[0].Id
                            });
                        }
                        else
                        {
                            NotUserInfo = NotUserInfo + "编号:" + exceldt.Rows[i]["编号"].ToString() + " 姓名:" + strName + " 会员卡类型:" + exceldt.Rows[i]["会员卡"].ToString() + " 剩余次数:" + exceldt.Rows[i]["剩余次数"].ToString() + "\n";
                            skipCount++;
                        }
                    }
                }
                int count = 0;
                using (_db)
                {

                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();
                    foreach (var item in list)
                    {
                        for (int i = 0; i < item.Balance; i++)
                        {
                            var NO = GetNO("S");
                            string sql = "INSERT INTO `AppCoupons` VALUES ('" + Guid.NewGuid() + "', now(), NULL, NULL, NULL, '" + NO + "', '36eddd1f-a7e8-570c-10eb-39f2e227c526', '成人减脂', 0, 3, now(), '2021-06-05 02:11:36.191000', '" + item.OwnerId + "', '0001-01-01 00:00:00.000000');";
                            await InsertAsync(_db, sql);
                        }
                        count++;
                    }
                }

                using (StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "成人减脂.txt"))))
                {
                    txtFile.WriteLine(NotUserInfo);
                }
                return "导入:" + count + " 跳出:" + skipCount;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        /// <summary>
        /// 大器械
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("dqx-card")]
        public async Task<string> DqxCard(IFormFile excelfile)
        {
            /*
             * TODO  
             * 1.读取表格数据
             * 2.查找名字
             * 3.插入数据
             */
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            DataTable dt = GetCrjzCardDT();
            var exceldt = GetExcelDT(excelfile, file, dt, 2);

            List<DisposeDBModel> list = new List<DisposeDBModel>();
            string NotUserInfo = null;
            try
            {

                int skipCount = 0;
                using (_db)
                {
                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();

                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        string strName = exceldt.Rows[i]["姓名"].ToString();

                        cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where `Name`='" + strName + "'";
                        var result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        if (!result.Any())
                        {
                            cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where Nickname='" + strName + "'";
                            result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        }
                        if (result.Any())
                        {
                            list.Add(new DisposeDBModel
                            {
                                Balance = decimal.Parse(exceldt.Rows[i]["剩余次数"].ToString()),
                                Discriminator = exceldt.Rows[i]["会员卡"].ToString(),
                                OwnerId = result[0].Id
                            });
                        }
                        else
                        {
                            NotUserInfo = NotUserInfo + "编号:" + exceldt.Rows[i]["编号"].ToString() + " 姓名:" + strName + " 会员卡类型:" + exceldt.Rows[i]["会员卡"].ToString() + " 剩余次数:" + exceldt.Rows[i]["剩余次数"].ToString() + "\n";
                            skipCount++;
                        }
                    }
                }
                int count = 0;
                using (_db)
                {

                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();
                    foreach (var item in list)
                    {
                        for (int i = 0; i < item.Balance; i++)
                        {
                            var NO = GetNO("S");
                            string sql = "INSERT INTO `AppCoupons` VALUES ('" + Guid.NewGuid() + "', now(), NULL, NULL, NULL, '" + NO + "', '28bf1964-6bda-edbe-f116-39f2e22bc1fa', '大器械', 0, 3, now(), '2021-06-05 02:11:36.191000', '" + item.OwnerId + "', '0001-01-01 00:00:00.000000');";
                            await InsertAsync(_db, sql);
                        }
                        count++;
                    }
                }

                using (StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "大器械.txt"))))
                {
                    txtFile.WriteLine(NotUserInfo);
                }
                return "导入:" + count + " 跳出:" + skipCount;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        /// <summary>
        /// 孕妇私教
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("yfsj-card")]
        public async Task<string> YfsjCard(IFormFile excelfile)
        {
            /*
             * TODO  
             * 1.读取表格数据
             * 2.查找名字
             * 3.插入数据
             */
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            DataTable dt = GetCrjzCardDT();
            var exceldt = GetExcelDT(excelfile, file, dt, 4);

            List<DisposeDBModel> list = new List<DisposeDBModel>();
            string NotUserInfo = null;
            try
            {

                int skipCount = 0;
                using (_db)
                {
                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();

                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        string strName = exceldt.Rows[i]["姓名"].ToString();

                        cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where `Name`='" + strName + "'";
                        var result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        if (!result.Any())
                        {
                            cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where Nickname='" + strName + "'";
                            result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        }
                        if (result.Any())
                        {
                            list.Add(new DisposeDBModel
                            {
                                Balance = decimal.Parse(exceldt.Rows[i]["剩余次数"].ToString()),
                                Discriminator = exceldt.Rows[i]["会员卡"].ToString(),
                                OwnerId = result[0].Id
                            });
                        }
                        else
                        {
                            NotUserInfo = NotUserInfo + "编号:" + exceldt.Rows[i]["编号"].ToString() + " 姓名:" + strName + " 会员卡类型:" + exceldt.Rows[i]["会员卡"].ToString() + " 剩余次数:" + exceldt.Rows[i]["剩余次数"].ToString() + "\n";
                            skipCount++;
                        }
                    }
                }
                int count = 0;
                using (_db)
                {

                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();
                    foreach (var item in list)
                    {
                        for (int i = 0; i < item.Balance; i++)
                        {
                            var NO = GetNO("S");
                            string sql = "INSERT INTO `AppCoupons` VALUES ('" + Guid.NewGuid() + "', now(), NULL, NULL, NULL, '" + NO + "', 'bd9103ff-4232-15e1-1854-39f3082650cb', '孕妇私教', 0, 3, now(), '2021-06-05 02:11:36.191000', '" + item.OwnerId + "', '0001-01-01 00:00:00.000000');";
                            await InsertAsync(_db, sql);
                        }
                        count++;
                    }
                }

                using (StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "孕妇私教.txt"))))
                {
                    txtFile.WriteLine(NotUserInfo);
                }
                return "导入:" + count + " 跳出:" + skipCount;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        /// <summary>
        /// 孕妇团教
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("yftj-card")]
        public async Task<string> YftjCard(IFormFile excelfile)
        {
            /*
             * TODO  
             * 1.读取表格数据
             * 2.查找名字
             * 3.插入数据
             */
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            DataTable dt = GetCrjzCardDT();
            var exceldt = GetExcelDT(excelfile, file, dt, 5);

            List<DisposeDBModel> list = new List<DisposeDBModel>();
            string NotUserInfo = null;
            try
            {

                int skipCount = 0;
                using (_db)
                {
                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();

                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        string strName = exceldt.Rows[i]["姓名"].ToString();

                        cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where `Name`='" + strName + "'";
                        var result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        if (!result.Any())
                        {
                            cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where Nickname='" + strName + "'";
                            result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        }
                        if (result.Any())
                        {
                            list.Add(new DisposeDBModel
                            {
                                Balance = decimal.Parse(exceldt.Rows[i]["剩余次数"].ToString()),
                                Discriminator = exceldt.Rows[i]["会员卡"].ToString(),
                                OwnerId = result[0].Id
                            });
                        }
                        else
                        {
                            NotUserInfo = NotUserInfo + "编号:" + exceldt.Rows[i]["编号"].ToString() + " 姓名:" + strName + " 会员卡类型:" + exceldt.Rows[i]["会员卡"].ToString() + " 剩余次数:" + exceldt.Rows[i]["剩余次数"].ToString() + "\n";
                            skipCount++;
                        }
                    }
                }
                int count = 0;
                using (_db)
                {

                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();
                    foreach (var item in list)
                    {
                        for (int i = 0; i < item.Balance; i++)
                        {
                            var NO = GetNO("S");
                            string sql = "INSERT INTO `AppCoupons` VALUES ('" + Guid.NewGuid() + "', now(), NULL, NULL, NULL, '" + NO + "', '508bdc91-e963-2e8e-0f49-39f30826c3df', '孕妇团教', 0, 3, now(), '2021-06-05 02:11:36.191000', '" + item.OwnerId + "', '0001-01-01 00:00:00.000000');";
                            await InsertAsync(_db, sql);
                        }
                        count++;
                    }
                }

                using (StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "孕妇团教.txt"))))
                {
                    txtFile.WriteLine(NotUserInfo);
                }
                return "导入:" + count + " 跳出:" + skipCount;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        /// <summary>
        /// 产后修复
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("chxf-card")]
        public async Task<string> ChxfCard(IFormFile excelfile)
        {
            /*
             * TODO  
             * 1.读取表格数据
             * 2.查找名字
             * 3.插入数据
             */
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.xlsx";
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            DataTable dt = GetCrjzCardDT();
            var exceldt = GetExcelDT(excelfile, file, dt, 6);

            List<DisposeDBModel> list = new List<DisposeDBModel>();
            string NotUserInfo = null;
            try
            {

                int skipCount = 0;
                using (_db)
                {
                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();

                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        string strName = exceldt.Rows[i]["姓名"].ToString();

                        cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where `Name`='" + strName + "'";
                        var result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        if (!result.Any())
                        {
                            cmd.CommandText = @"SELECT `Id`,`Name`,Nickname FROM `AbpUsers` where Nickname='" + strName + "'";
                            result = await ReadAllAsync(await cmd.ExecuteReaderAsync());
                        }
                        if (result.Any())
                        {
                            list.Add(new DisposeDBModel
                            {
                                Balance = decimal.Parse(exceldt.Rows[i]["剩余次数"].ToString()),
                                Discriminator = exceldt.Rows[i]["会员卡"].ToString(),
                                OwnerId = result[0].Id
                            });
                        }
                        else
                        {
                            NotUserInfo = NotUserInfo + "编号:" + exceldt.Rows[i]["编号"].ToString() + " 姓名:" + strName + " 会员卡类型:" + exceldt.Rows[i]["会员卡"].ToString() + " 剩余次数:" + exceldt.Rows[i]["剩余次数"].ToString() + "\n";
                            skipCount++;
                        }
                    }
                }
                int count = 0;
                using (_db)
                {

                    await _db.Connection.OpenAsync();
                    var cmd = _db.Connection.CreateCommand();
                    foreach (var item in list)
                    {
                        for (int i = 0; i < item.Balance; i++)
                        {
                            var NO = GetNO("S");
                            string sql = "INSERT INTO `AppCoupons` VALUES ('" + Guid.NewGuid() + "', now(), NULL, NULL, NULL, '" + NO + "', 'e88e16a2-f575-04f9-850d-39f308270fe4', '产后修复', 0, 3, now(), '2021-06-05 02:11:36.191000', '" + item.OwnerId + "', '0001-01-01 00:00:00.000000');";
                            await InsertAsync(_db, sql);
                        }
                        count++;
                    }
                }

                using (StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "产后修复.txt"))))
                {
                    txtFile.WriteLine(NotUserInfo);
                }
                return "导入:" + count + " 跳出:" + skipCount;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        /// <summary>
        /// 余额数据同步
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("ds-savings")]
        public async Task DSSavings(IFormFile excelfile)
        {
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.json";
            var path = Path.Combine(sWebRootFolder, sFileName);
            FileInfo file = new FileInfo(path);

            using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
            {
                excelfile.CopyTo(fs);
                fs.Flush();
            }
            string NotUserInfo = null;
            string package = System.IO.File.ReadAllText(path, Encoding.Default);

            List<SavingModel> list = JsonConvert.DeserializeObject<List<SavingModel>>(package);

            foreach (var item in list)
            {
                var handler = new HttpClientHandler() { AutomaticDecompression = DecompressionMethods.GZip };
                UserInfoModel userInfo = null;
                using (var httpClient = new HttpClient())
                {
                    var requestUri = "https://loosen.microfeel.net/api/user/getList?Name=" + item.Name;
                    var httpResponseMessage = await httpClient.GetAsync(requestUri);
                    var responseConetent = await httpResponseMessage.Content.ReadAsStringAsync();
                    userInfo = JsonConvert.DeserializeObject<List<UserInfoModel>>(responseConetent).FirstOrDefault();
                }
                if (userInfo == null)
                {
                    NotUserInfo += item.Name + "无用户信息" + "\n";
                    continue;
                }
                //WithholdInputModel withhold = null;
                switch (item.Type)
                {
                    case SavingType.通用余额:
                        NotUserInfo += await NewMethod(item, handler, userInfo, userInfo.Balance_Common) + "\n";

                        break;
                    case SavingType.团教余额:
                        NotUserInfo += await NewMethod(item, handler, userInfo, userInfo.Balance_Group) + "\n";
                        break;
                    case SavingType.私教余额:
                        NotUserInfo += await NewMethod(item, handler, userInfo, userInfo.Balance_Personal) + "\n";
                        break;
                    default:
                        break;
                }

            }
            using StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "余额3月.txt")));
            txtFile.WriteLine(NotUserInfo);


            //return true;
        }

        /// <summary>
        /// 卡卷数据同步
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("ds-coupons")]
        public async Task DsCoupons(IFormFile excelfile)
        {
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{Guid.NewGuid()}.json";
            var path = Path.Combine(sWebRootFolder, sFileName);
            FileInfo file = new FileInfo(path);

            using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
            {
                excelfile.CopyTo(fs);
                fs.Flush();
            }
            string NotUserInfo = null;
            string package = System.IO.File.ReadAllText(path, Encoding.Default);

            List<CouponModel> list = JsonConvert.DeserializeObject<List<CouponModel>>(package);

            foreach (var item in list)
            {
                var handler = new HttpClientHandler() { AutomaticDecompression = DecompressionMethods.GZip };
                UserInfoModel userInfo = null;
                using (var httpClient = new HttpClient())
                {
                    var requestUri = "https://loosen.microfeel.net/api/user/getList?Name=" + item.Name;
                    var httpResponseMessage = await httpClient.GetAsync(requestUri);
                    var responseConetent = await httpResponseMessage.Content.ReadAsStringAsync();
                    userInfo = JsonConvert.DeserializeObject<List<UserInfoModel>>(responseConetent).FirstOrDefault();
                }
                if (userInfo == null)
                {
                    NotUserInfo += item.Name + "无用户信息" + "\n";
                    continue;
                }
                List<Coupon> coupons = null;
                using (var httpClient = new HttpClient())
                {
                    var requestUri = "https://loosen.microfeel.net/api/user/get-coupon?UserId=" + userInfo.Id + "&CouponTypeName=" + item.Type;
                    var httpResponseMessage = await httpClient.GetAsync(requestUri);
                    var responseConetent = await httpResponseMessage.Content.ReadAsStringAsync();
                    coupons = JsonConvert.DeserializeObject<List<Coupon>>(responseConetent);
                }
                if (coupons == null)
                {
                    NotUserInfo += item.Name + "无卡卷信息" + "\n";
                    continue;
                }
                if (coupons.Count > item.Count)
                {
                    int useCount = coupons.Count - item.Count;
                    for (int i = 0; i < useCount; i++)
                    {
                        using var httpClient = new HttpClient();
                        var requestUri = "https://loosen.microfeel.net/api/user/use-coupon?couponId=" + coupons[i].Id;
                        var httpResponseMessage = await httpClient.GetAsync(requestUri);
                        var responseConetent = await httpResponseMessage.Content.ReadAsStringAsync();
                        httpResponseMessage.EnsureSuccessStatusCode();
                        if (httpResponseMessage.IsSuccessStatusCode)
                        {
                            NotUserInfo += "消耗:" + coupons[i].CouponTypeName + "\n";
                        }
                    }
                }
            }
            using StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, "卡卷.txt")));
            txtFile.WriteLine(NotUserInfo);
        }

        /// <summary>
        /// 私教产品数据同步
        /// </summary>
        /// <param name="excelfile"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("ds-personal-course-product")]
        public async Task DsPersonalCourseProduct(IFormFile excelfile)
        {
            string sWebRootFolder = _webHostEnvironment.WebRootPath;
            string sFileName = $"{GetNO("P")}.json";
            var path = Path.Combine(sWebRootFolder, sFileName);

            string package = ReadTest(excelfile, path);
            string NotUserInfo = null;

            List<ProductModel> list = JsonConvert.DeserializeObject<List<ProductModel>>(package);
            foreach (var item in list)
            {
                var handler = new HttpClientHandler() { AutomaticDecompression = DecompressionMethods.GZip };
                using var http = new HttpClient(handler);
                var content = new StringContent(JsonConvert.SerializeObject(item), Encoding.UTF8, "application/json");


                var response = await http.PostAsync("https://loosen.microfeel.net/api/product/add/personal-course-product", content);

                response.EnsureSuccessStatusCode();

                if (response.IsSuccessStatusCode)
                {
                    NotUserInfo += item.UserName + "的" + item.Price + item.CourseName + "添加成功\n";
                }
            }
            using StreamWriter txtFile = new StreamWriter(System.IO.File.Create(Path.Combine(sWebRootFolder, $"{GetNO("卡卷")}.txt")));
            txtFile.WriteLine(NotUserInfo);
        }

        private static string ReadTest(IFormFile excelfile, string path)
        {
            FileInfo file = new FileInfo(path);
            using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
            {
                excelfile.CopyTo(fs);
                fs.Flush();
            }
            string package = System.IO.File.ReadAllText(path, Encoding.Default);
            return package;
        }

        public class Coupon
        {
            public string SerialNO { get; set; }

            public Guid CouponTypeId { get; set; }

            public string CouponTypeName { get; set; }

            public Guid Id { get; set; }

            public ChannelType ChannelType { get; set; }

            public DateTime StartValidTime { get; set; }

            public DateTime EndValidTime { get; set; }

            public virtual Guid UseUserId { get; set; }

            //public virtual AppUser User { get; set; }

            /// <summary>
            /// 失效日期
            /// </summary>
            public DateTime LoseEfficacyTime { get; set; }
        }

        private static async Task<string> NewMethod(SavingModel item, HttpClientHandler handler, UserInfoModel userInfo, decimal balance)
        {
            string NotUserInfo = null;
            if (balance == item.Price)
            {
                NotUserInfo += item.Name + item.Type.ToString() + "信息准确，";

            }
            else if (balance < item.Price)
            {
                RechargeInputModel RechargeInput = new RechargeInputModel
                {
                    Balance = item.Price - balance,
                    ChannelType = ChannelType.数据导入,
                    Money = item.Price - balance,
                    OwnerId = userInfo.Id,
                    SavingSource = SavingSource.平台已有数据导入,
                    SavingType = item.Type
                };
                NotUserInfo += item.Name + "充值" + await PostRecharge(handler, RechargeInput);
            }
            else if (balance > item.Price)
            {
                WithholdInputModel withhold = new WithholdInputModel
                {
                    PayPrice = balance - item.Price,
                    PayType = PayType.余额支付,
                    UserId = userInfo.Id
                };
                NotUserInfo += item.Name + "扣款" + await PostWithhold(handler, withhold);

            }

            return NotUserInfo;
        }

        private static async Task<string> PostWithhold(HttpClientHandler handler, WithholdInputModel withholds)
        {
            string NotUserInfo = "";
            using (var http = new HttpClient(handler))
            {

                var content = new StringContent(JsonConvert.SerializeObject(withholds), Encoding.UTF8, "application/json");


                var response = await http.PostAsync("https://loosen.microfeel.net/api/saving/withhold", content);

                response.EnsureSuccessStatusCode();

                if (response.IsSuccessStatusCode)
                {
                    NotUserInfo += withholds.PayType.ToString() + ":" + withholds.PayPrice;
                }

            }

            return NotUserInfo;
        }

        private static async Task<string> PostRecharge(HttpClientHandler handler, RechargeInputModel recharge)
        {
            string NotUserInfo = "";
            using (var http = new HttpClient(handler))
            {

                var content = new StringContent(JsonConvert.SerializeObject(recharge), Encoding.UTF8, "application/json");


                var response = await http.PostAsync("https://loosen.microfeel.net/api/saving/recharge", content);

                response.EnsureSuccessStatusCode();

                if (response.IsSuccessStatusCode)
                {
                    NotUserInfo += recharge.SavingType.ToString() + ":" + recharge.Balance;
                }

            }

            return NotUserInfo;
        }

        private static short _sn = 0;
        private static readonly object Locker = new object();

        /// <summary>
        /// 生成编号
        /// </summary>
        /// <param name="sign"></param>
        /// <returns></returns>
        public static string GetNO(string sign)
        {
            lock (Locker)  //lock 关键字可确保当一个线程位于代码的临界区时，另一个线程不会进入该临界区。 
            {
                if (_sn == short.MaxValue)
                {
                    _sn = 0;
                }
                else
                {
                    _sn++;
                }

                return sign + DateTime.Now.ToString("yyyyMMddHHmmss") + (_sn.ToString().PadLeft(5, '0'));
            }
        }
        private async Task InsertAsync(AppDb db, DisposeDBModel disposeDBModel)
        {
            var cmd = db.Connection.CreateCommand() as MySqlCommand;
            cmd.CommandText = @"INSERT INTO `AppSavings` VALUES ('" + Guid.NewGuid() + "', '{}', NULL, now(), NULL, NULL, NULL, b'0', NULL, NULL, " + disposeDBModel.Balance + ", " + disposeDBModel.Balance + ", 3, '" + disposeDBModel.OwnerId + "', '" + disposeDBModel.Discriminator + "', NULL);";
            await cmd.ExecuteNonQueryAsync();
            //int Id = (int)cmd.LastInsertedId;
        }

        private async Task InsertAsync(AppDb db, string sql)
        {
            var cmd = db.Connection.CreateCommand() as MySqlCommand;
            cmd.CommandText = sql;
            await cmd.ExecuteNonQueryAsync();
            //int Id = (int)cmd.LastInsertedId;
        }

        private async Task<List<UserModel>> ReadAllAsync(DbDataReader reader)
        {
            var posts = new List<UserModel>();
            using (reader)
            {
                while (await reader.ReadAsync())
                {
                    UserModel post = new UserModel
                    {
                        Id = await reader.GetFieldValueAsync<Guid>(0),
                        Name = await reader.GetFieldValueAsync<string>(1),
                        Nickname = await reader.GetFieldValueAsync<string>(2)
                    };
                    posts.Add(post);
                }
            }
            return posts;
        }

        public class UserModel
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public string Nickname { get; set; }
        }

        public class DisposeDBModel
        {
            public Guid OwnerId { get; set; }
            public decimal Balance { get; set; }
            public string Discriminator { get; set; }
        }

        private static DataTable GetExcelDT(IFormFile excelfile, FileInfo file, DataTable dt, int index = 0)
        {
            try
            {
                using (FileStream fs = new FileStream(file.ToString(), FileMode.Create))
                {
                    excelfile.CopyTo(fs);
                    fs.Flush();
                }



                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[index];
                    int startRowIndx = sheet.Dimension.Start.Row + (true ? 1 : 0);
                    for (int r = startRowIndx; r <= sheet.Dimension.End.Row; r++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int c = sheet.Dimension.Start.Column; c <= sheet.Dimension.End.Column; c++)
                        {
                            if (dr.ItemArray.Length < c)
                            {
                                continue;
                            }
                            dr[c - 1] = (sheet.Cells[r, c].Value ?? DBNull.Value);
                        }
                        dt.Rows.Add(dr);
                    }


                }

                return dt;
            }
            catch (Exception ex)
            {
                throw new NotImplementedException(ex.Message);
            }
        }

        private static DataTable GetPetCardDT()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("编号", Type.GetType("System.String"));
            dt.Columns.Add("姓名", Type.GetType("System.String"));
            dt.Columns.Add("会员卡", Type.GetType("System.String"));
            dt.Columns.Add("总计", Type.GetType("System.String"));
            dt.Columns.Add("12月剩余金额", Type.GetType("System.String"));
            dt.Columns.Add("1月存入金额", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/1", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/2", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/3", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/4", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/5", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/6", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/7", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/8", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/9", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/10", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/11", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/12", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/13", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/14", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/15", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/16", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/17", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/18", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/19", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/20", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/21", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/22", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/23", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/24", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/25", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/26", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/27", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/28", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/29", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/30", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/31", Type.GetType("System.String"));
            dt.Columns.Add("1月累计消费额", Type.GetType("System.String"));
            dt.Columns.Add("剩余金额", Type.GetType("System.String"));
            return dt;
        }

        private static DataTable GetCrjzCardDT()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("编号", Type.GetType("System.String"));
            dt.Columns.Add("姓名", Type.GetType("System.String"));
            dt.Columns.Add("会员卡", Type.GetType("System.String"));
            dt.Columns.Add("总计", Type.GetType("System.String"));
            dt.Columns.Add("单价", Type.GetType("System.String"));
            dt.Columns.Add("12月剩余次数", Type.GetType("System.String"));
            dt.Columns.Add("12月剩余金额", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/1", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/2", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/3", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/4", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/5", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/6", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/7", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/8", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/9", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/10", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/11", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/12", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/13", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/14", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/15", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/16", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/17", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/18", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/19", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/20", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/21", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/22", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/23", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/24", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/25", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/26", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/27", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/28", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/29", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/30", Type.GetType("System.String"));
            dt.Columns.Add("2020/1/31", Type.GetType("System.String"));
            dt.Columns.Add("1月上课次数", Type.GetType("System.String"));
            dt.Columns.Add("剩余次数", Type.GetType("System.String"));
            dt.Columns.Add("剩余金额", Type.GetType("System.String"));
            return dt;
        }
    }

    public class RechargeInputModel
    {
        [Required]
        public decimal Balance { get; set; }

        [Required]
        public decimal Money { get; set; }

        public DateTime? ExpiryDate { get; set; }

        public SavingSource SavingSource { get; set; }

        [Required]
        public Guid OwnerId { get; set; }

        [Required]
        public SavingType SavingType { get; set; }

        [Required]
        public ChannelType ChannelType { get; set; }
    }

    public enum SavingSource
    {
        充值,
        赠送,
        退款,
        平台已有数据导入
    }

    public enum SavingType
    {
        通用余额,
        团教余额,
        私教余额
    }
    public enum ChannelType
    {
        后台录入,
        用户操作,
        数据导入
    }



    internal class SavingModel
    {
        public string Name { get; set; }

        public SavingType Type { get; set; }

        public decimal Price { get; set; }
    }

    internal class CouponModel
    {
        public string Name { get; set; }

        public string Type { get; set; }

        public int Count { get; set; }
    }

    internal class ProductModel
    {
        public string CourseName { get; set; }

        public string UserName { get; set; }

        public decimal Price { get; set; }

        public int PeopleNumber { get; set; }

        public string Label { get; set; }
    }

    internal class UserInfoModel
    {
        public Guid Id { get; set; }

        /// <summary>
        /// 是否是Vip
        /// </summary>
        public bool IsVip { get; set; }

        /// <summary>
        /// 会员剩余天数
        /// </summary>
        public int Surplus { get; set; }

        /// <summary>
        /// 通用余额
        /// </summary>
        public decimal Balance_Common { get; set; }

        /// <summary>
        /// 团教课余额
        /// </summary>
        public decimal Balance_Group { get; set; }

        /// <summary>
        /// 私教课余额
        /// </summary>
        public decimal Balance_Personal { get; set; }

        /// <summary>
        /// 卡卷剩余数
        /// </summary>
        public int Card_detail { get; set; }

        /// <summary>
        /// 存在天数
        /// </summary>
        public int Exist { get; set; }

        public string Name { get; set; }

        public string PhoneNumber { get; set; }

        public string Icon { get; set; }

        public string Nickname { get; set; }

        /// <summary>
        /// 是否需要补充
        /// </summary>
        public bool IsSupplement { get; set; }

        /// <summary>
        /// 会员版本名称
        /// </summary>
        public string MemberEditionName { get; set; }

        /// <summary>
        /// 特权Url
        /// </summary>
        public string EditionUrl { get; set; }

        /// <summary>
        /// 版本图片
        /// </summary>
        public string EditionImg { get; set; }

        public string BackgroundImage { get; set; }
    }

    public enum PayType
    {
        微信支付,
        余额支付,
        卡卷支付,
        团课余额支付,
        私教课余额支付
    }

    public class WithholdInputModel
    {
        public decimal PayPrice { get; set; }

        public Guid UserId { get; set; }

        public PayType PayType { get; set; }
    }
}
