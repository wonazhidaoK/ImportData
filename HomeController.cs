using ImportData.Base;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
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
            var exceldt = GetExcelDT(excelfile, file, dt,1);

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
            int Id = (int)cmd.LastInsertedId;
        }

        private async Task InsertAsync(AppDb db, string sql)
        {
            var cmd = db.Connection.CreateCommand() as MySqlCommand;
            cmd.CommandText = sql;
            await cmd.ExecuteNonQueryAsync();
            int Id = (int)cmd.LastInsertedId;
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
                            if(dr.ItemArray.Length< c)
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
}
