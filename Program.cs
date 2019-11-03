using System;
using PuppeteerSharp;
using System.Threading.Tasks;
using System.Windows;
using System.Linq;

namespace scrapperc_
{
    class Program
    {
         public static async Task Main(string[] args)
        {
		
            Console.WriteLine("Downloading chromium");
            //await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision);
            Console.WriteLine("Navigating to google.com");
            string skill = "java", location = "nj", length = "10", minExp = "5";
            //string workPermits = "us citizenship|green card|have h1|employment auth document|need h1|tn permit holder";
            string workPermits = "";
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
            //workPermits=us%20citizenship%7Cgreen%20card%7Chave%20h1%7Cemployment%20auth%20document%7Cneed%20h1%7Ctn%20permit%20holder;
            //%20 = space; %7C = |
            //workPermits=us citizenship|green card|have h1|employment auth document|need h1|tn permit holder;

            // Download the Chromium revision if it does not already exist
            var options = new LaunchOptions{Headless=true, ExecutablePath = "C:\\Users\\Shazad\\Desktop\\scrapperc#\\.local-chromium\\Win64-686378\\chrome-win\\chrome.exe"};
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
                await page.TypeAsync("#username","imran@prosistech.com");
                await page.TypeAsync("#password","Prosis2019");
                await page.ClickAsync(".sign-in-btn");
                string class1 = ".search-result";
                try{
                    await page.WaitForSelectorAsync(class1);
                }catch{
                    Console.Write("Kuch to garbar he\nExiting 1...\n");
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
                
                for(int i = 0;i<urls.Length;i++)
                {
                    string[] temp = urls[i].Split('$');
                    await page.GoToAsync(temp[1],1500000000);
                    string class2 = ".media-body";
                    try{
                        await page.WaitForSelectorAsync(class2);
                    }catch{
                        Console.Write("Kuch to garbar he\nExiting 2...\n");
                        return;
                    }
                    var jsSelectAllInfo = @"Array.from(document.querySelectorAll('.media-body')).map(x => x.innerText);";
                    //var jsSelectEmail = @"document.querySelectorAll('.media-body')[2].innerText;";
                    
                    var allInfo = await page.EvaluateExpressionAsync<string[]>(jsSelectAllInfo);
                    Console.WriteLine((i+1)+"- "+temp[0]+":");
                    for(int j = 1;j < allInfo.Length;j++)
                    {
                        if(allInfo[j].Contains("Dice Data"))
                        {
                            break;
                        }
                        Console.WriteLine("\t"+allInfo[j]);
                    }
			Console.WriteLine("test");
                }
                

            }
                //System.IO.File.WriteAllText(@"C:\\Users\\Compro\\Documents\\Dice\\Dice\\page1.txt",content);
                

                //search for persons
                /*string url = "https://www.dice.com/employer/talent/search;q="+skill+";location="+location+";page=1;pageSize=20;sortBy=relevance;sortByDirection=desc;excludeRecruiters=true;distance=50;distanceUoM=miles;lastActive=90;profileSource=resumes";
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
        }
    }
}
