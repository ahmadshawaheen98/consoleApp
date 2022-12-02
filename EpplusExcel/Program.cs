using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
namespace EpplusExcel
{
    public class person
    {
        
        public int id { set; get; }
        public string Name { set; get; }

        public string email { set; get; }

        public string isActive { set; get; }

    }
    public class context : DbContext
    {
        public DbSet<person> people { set; get; }
   


        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("data source=localhost; Initial catalog=UserExcal; Integrated security=true ");

            base.OnConfiguring(optionsBuilder);
        }
    }
    class Program
    {
       // private static object strarray;

        static void Main(string[] args)
        {
           Console.WriteLine("Welcome");
            string file =Console.ReadLine();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
               // ExcelPackage.LicenseContext = LicenseContext.Commercial;
               ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["data"];
                var persons = new Program().GetList<person>(sheet);
            }
        }
        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> ts = new List<T>();
            var columninfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
            new { index = n, columnName=sheet.Cells[1,n].Value.ToString()}
            );
       
                
            for (int row=2; row<=sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columninfo.SingleOrDefault(c => c.columnName == prop.Name).index;
                    var val = sheet.Cells[row, col].Value;
                    var proptype = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, proptype));

                   
                }

                context context = new context();
                context.Add(obj);
                context.SaveChanges();
                //ts.Add(obj);
            }
        
            return ts;
            
        }
    }
   
}
