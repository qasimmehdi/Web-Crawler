using System;
using System.Collections.Generic; 
using PuppeteerSharp;
using System.IO;
using System.Text;  
using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;
using System.Linq;

namespace scrapperc_
{
    class Program
    {
         public static async Task Main(string[] args)
        {
            /* Dependencies: Nudget EPPlus.core, PuppeteerSharp */
		
            
            Console.WriteLine("Starting");
            string skill = "java", location = "nj", length = "10", minExp = "5";
            //string workPermits = "us citizenship|green card|have h1|employment auth document|need h1|tn permit holder";
            string workPermits = "",username = "",password = "";
            Console.WriteLine("Email: ");
            username = Console.ReadLine();
            Console.WriteLine("Password: ");
            password = Console.ReadLine();
            Console.WriteLine("Skill: ");
            skill = Console.ReadLine();
            Console.WriteLine("Location: ");
            location = Console.ReadLine();
            Console.WriteLine("Length: ");
            length = Console.ReadLine();
            Console.WriteLine("Min Experience: ");
            minExp = Console.ReadLine();
            Console.WriteLine("Work Permit options:\nus citizenship,green card,have h1,employment auth document,need h1,tn permit holder\nLeave empty for no options");
            Console.WriteLine("To use multiple work permit options add | b/w them\ne.g. us citizenship|green card");
            Console.WriteLine("Work Permits: ");
            workPermits = Console.ReadLine();
            string[] names = new string[int.Parse(length)];
            string[] emails = new string[int.Parse(length)];
            string[] phones = new string[int.Parse(length)];
            
            //workPermits=us%20citizenship%7Cgreen%20card%7Chave%20h1%7Cemployment%20auth%20document%7Cneed%20h1%7Ctn%20permit%20holder;
            //%20 = space; %7C = |
            //workPermits=us citizenship|green card|have h1|employment auth document|need h1|tn permit holder;

            // Download the Chromium revision if it does not already exist
            
            try{
            new LaunchOptions{Headless=true, ExecutablePath = Directory.GetCurrentDirectory()+"\\.local-chromium\\Win64-686378\\chrome-win\\chrome.exe"};
            }
            catch{
             Console.WriteLine("Downloading chromium");
            await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision);
            }
            var options = new LaunchOptions{Headless=true, ExecutablePath = Directory.GetCurrentDirectory()+"\\.local-chromium\\Win64-686378\\chrome-win\\chrome.exe"};
            //Console.WriteLine("Downloading chromium");
            //var executablePath = new BrowserFetcher.GetExecutablePath(BrowserFetcher.DefaultRevision);
            Console.WriteLine("Launching headless chrome");
            using (var browser = await Puppeteer.LaunchAsync(options))
            {
            using (var page = await browser.NewPageAsync())
            {
                Console.WriteLine("Loading");
                string login_search = $"https://www.dice.com/employer/talent/search;q={skill};location={location};page=1;pageSize={length};sortBy=relevance;sortByDirection=desc;workPermits={workPermits};yOEMin={minExp};yOEMax=null;excludeRecruiters=true;contactInfo=email;distance=50;distanceUoM=miles;lastActive=90;profileSource=resumes";
                Console.WriteLine(login_search);
                await page.GoToAsync(login_search,1500000000);
                Console.WriteLine("Logging in");
                await page.TypeAsync("#username",username);
                await page.TypeAsync("#password",password);
                await page.ClickAsync(".sign-in-btn");
                string class1 = ".search-result";
                try{
                    await page.WaitForSelectorAsync(class1);
                }catch{
                    Console.Write("Error: Please Enter Correct Details or Restart The Program\nExiting 1...\n");
                    Console.ReadKey();
                    return;
                }
                string content = await page.GetContentAsync();
                var jsSelectAllAnchors = @"Array.from(document.querySelectorAll('.view-link')).map(a => a.innerText+'$'+a.href);";
                var urls = await page.EvaluateExpressionAsync<string[]>(jsSelectAllAnchors);
                /*foreach (string url in urls)
                {
                    Console.WriteLine($"Url: {url}");
                }*/
                //string url2 = "https://www.dice.com/employer/talent//profile/378f4c2e3d8b59538f65d7a4992ce34133c780cc?searchId=18a237c4-16db-4c30-8743-52cb01873326";
                int ab = 0;
                for(int i = 0;i<urls.Length;i++)
                {
                    
                    string[] temp = urls[i].Split('$');
                    await page.GoToAsync(temp[1],1500000000);
                    string class2 = ".media-body";
                    try{
                        await page.WaitForSelectorAsync(class2);
                    }catch{
                    Console.Write("Error: Please Check your Internet or Restart the Program\nExiting 2...\n");
                    Console.ReadKey();
                    }
                    var jsSelectAllInfo = @"Array.from(document.querySelectorAll('.media-body')).map(x => x.innerText);";
                    //var jsSelectEmail = @"document.querySelectorAll('.media-body')[2].innerText;";
                    
                    var allInfo = await page.EvaluateExpressionAsync<string[]>(jsSelectAllInfo);
                    names[i] = temp[0];
                    for(int j = 1,k=0;j < allInfo.Length;j++,k++)
                    {
                        if(allInfo[j].Contains("Dice Data"))
                        {
                            break;
                        }
                        if(j == 1){emails[ab] = allInfo[j];}
                        if(j >= 2){phones[ab] = allInfo[j];} 
                    }
                    ab++;
                    
                            Console.WriteLine(names[i]+" "+emails[i]+" "+phones[i]+ "\n" );
                    
                        
                    
                }
                

            }
             
                //System.IO.File.WriteAllText(@"C:\\Users\\Compro\\Documents\\Dice\\Dice\\page1.txt",content);
                

                //search for persons
                /*string url = "https://www.dice.com/employer/talent/search;q="+ski5
                ll+";location="+location+";page=1;pageSize=20;sortBy=relevance;sortByDirection=desc;excludeRecruiters=true;distance=50;distanceUoM=miles;lastActive=90;profileSource=resumes";
                await page.GoToAsync(url,WaitUntilNavigation.Networkidle0);
                
                content = await page.GetContentAsync();
                System.IO.File.WriteAllText(@"C:\\Users\\Compro\\Documents\\Dice\\Dice\\page2.txt",content);
                /*var jsSelectAllAnchors = @"Array.from(document.querySelectorAll('a')).map(a => a.href);";
                var urls = await page.EvaluateExpressionAsync<string[]>(jsSelectAllAnchors);
                foreach (string url in urls)
                {
                    Console.WriteLine($"Url: {url}");
                }
            }
            //Console.WriteLine("Hello World!");*/
            }
            new Program().excl(names,emails,phones,skill);
        }
        public void excl(string[] name,string[] email,string[] phone,string filenaam){

            string rootFolder = Directory.GetCurrentDirectory();  
            Random rnd = new Random();
            string fileName = $"{filenaam+rnd.Next(1,100).ToString()}.xlsx";  
  
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
     			using (ExcelPackage package = new ExcelPackage(file))  
            {  
  
                
  
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Dice");  
                int totalRows = name.Length;  

                worksheet.Cells[1, 1].Value = "Name";  
                worksheet.Cells[1, 2].Value = "Email";  
                worksheet.Cells[1, 3].Value = "Phone";   
                  
                for (int row = 2,i=0; row <= totalRows + 1; row++,i++)  
                {  
                    worksheet.Cells[row, 1].Value = name[i];  
                    worksheet.Cells[row, 2].Value = email[i];  
                    worksheet.Cells[row, 3].Value = phone[i];  
                     
                }  
  
                package.Save(); 
                Console.WriteLine($"File:{fileName} Creation successfull ");
  
            }  

        }
    }
}
