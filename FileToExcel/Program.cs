using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileToExcel
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string path = "note1.txt";
            string text = "aram,ababa,ahaha";
            string m;
            int e = 0, f = 0, i = 0, j = 1; ;
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                await writer.WriteLineAsync(text);
                writer.Close();
            }
            using (StreamReader reader = new StreamReader(path))
            {
                m = reader.ReadToEnd();
                reader.Close();
            }
            
            


            try 
            {
                using (var helper=new ExcelHelper())
                {
                    if(helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "test.xlsx")))
                    {
                        for (e = 0, f = 0, i = 0; i < m.Length; i++)
                        {
                            if (m[i] == ',')
                            {
                                f = i;
                                helper.Set("A", j, m.Substring(e, f - e-1));
                                j++;
                            }
                            e = f+1;
                        }
                        
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
