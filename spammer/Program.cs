using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spammer
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.BufferHeight = 30000;

            Matcher matcher = new Matcher();
            matcher.Match();

            Console.Read();
        }
    }

    public class MailSender
    {
        Person[] ccs;
        Person[] bccs;
        ApplicationCreator applicationCreator = new ApplicationCreator();
        ExcelScanner projectScanner = new ExcelScanner();
        Outlook.Application outlookApp;
        List<Project> projects;
        Person sender = new Person("Jackson, Michael", "michael.jackson");
        Person supervisor1 = new Person("John, Elton", "elton.john");
        Person supervisor2 = new Person("Starr, Ringo", "ringo.starr");
        Person interested1 = new Person("Page, Jimmy", "jimmy.page");
        Person interested2 = new Person("Yorke, Thom", "thom.yorke");

        public MailSender()
        {
            outlookApp = applicationCreator.GetOutlookApplication();
            projects = projectScanner.RetrieveMatches();
            ccs = new Person[] { supervisor1, supervisor2, interested1 };
            bccs = new Person[] { interested2 };
        }

        public MailSender(Person s, Person[] c, Person[] b)
        {
            sender = s;
            ccs = c;
            bccs = b;

            outlookApp = applicationCreator.GetOutlookApplication();
            projects = projectScanner.RetrieveMatches();
        }

        public void SendMail(Project project)
        {
            Outlook.MailItem mail;
            string shouldSend;
            TextOperator textOperator = new TextOperator();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Sending mail " + ", ID " + project.id + "...");
            Console.ResetColor();

            mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;

            mail.To = BuildTo(new Person[] { project.contact1, project.contact2 });
            mail.CC = BuildTo(ccs);
            mail.BCC = BuildTo(bccs);

            mail.Subject = BuildSubject(project.id, project.title, project.city);
            mail.HTMLBody = BuildBody(project.contact1, project.contact2, project.title, project.candidates, mail);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\nEmail to be sent:");
            Console.ResetColor();

            Console.WriteLine("To: " + mail.To);
            Console.WriteLine("\nCC: " + mail.CC);
            Console.WriteLine("\nBCC: " + mail.BCC);
            Console.WriteLine("\nSubject: " + mail.Subject);
            Console.WriteLine("\nBody: " + mail.HTMLBody);

            Console.WriteLine("\nSend? (y/n)");
            shouldSend = Console.ReadLine();

            if (shouldSend == "y")
            {
                mail.Send();
                textOperator.WriteSentProject(project);
            }
        }

        private string BuildTo(string[] addresses)
        {
            string to = "";

            for (int i = 0; i < addresses.Length; i++)
                to += addresses[i] + ";";

            return to;
        }

        private string BuildTo(Person[] people)
        {
            List<string> addresses = new List<string>();
            IdFinder idFinder;

            for (int i = 0; i < people.Length; i++)
            {
                if (people[i].email == null)
                {
                    idFinder = new IdFinder(people[i]);
                    people[i].SetId(idFinder.FindId());
                }

                addresses.Add(people[i].email);
            }

            string to = "";

            for (int i = 0; i < addresses.Count; i++)
                to += addresses[i] + ";";

            return to;
        }

        private string BuildSubject(string id, string title, string city)
        {
            string subject = "Job Application - " + id + " " + title + " - " + city;

            return subject;
        }

        private string BuildBody(Person contact1, Person contact2, string title, List<Person> candidates, Outlook.MailItem mail)
        {
            string imageCid;
            string body = "<body>Hi ";

            if (contact1.email != null)
                body += contact1.nameFirst + ", ";
            if (contact2.email != null)
                body += contact2.nameFirst + ",";

            body += "<br /><br />Hello!<br /><br />My name is " +
                sender.nameFirst + " " + sender.nameLast +
                "I see an open position for a " +
                title +
                ". ";

            if (candidates.Count == 1)
                body += "I'd like to recommend this person for the job.";
            else
                body += "I'd like to recommend these persons for the job.";

            body += "<br /><br />Attached is an overview of their skills.<br /><br />";

            for (int i = 0; i < candidates.Count; i++)
            {
                imageCid = candidates[i].id + ".bmp@123";
                Outlook.Attachment attachment = mail.Attachments.Add(Directory.GetCurrentDirectory() + "\\" + candidates[i].id + ".png", Outlook.OlAttachmentType.olEmbeddeditem, null, null);
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid);
                body += "<img src=\"cid:" + imageCid + "\" /><br /><br />";
            }

            body += "Thanks,<br /><br />" +
                sender.nameFirst + " " + sender.nameLast;

            return body;
        }
    }

    public class IdFinder
    {
        Outlook.Application outlookApp;
        Outlook.AddressList globalAddressList;
        Person person;

        public IdFinder(Person p)
        {
            ApplicationCreator applicationCreator = new ApplicationCreator();
            outlookApp = applicationCreator.GetOutlookApplication();

            globalAddressList = outlookApp.Session.GetGlobalAddressList();
            person = p;
        }

        public string FindId()
        {
            string id = "";
            List<string> ids = new List<string>();
            List<Person> matches = new List<Person>();
            int index = FindLastNameIndex(1, globalAddressList.AddressEntries.Count);
            int counter = 1;
            Outlook.ExchangeUser exUser;
            Person temp;

            if (person.name != null)
                for (int i = 0; i < 2; i++)
                {
                    exUser = globalAddressList.AddressEntries[index + counter].GetExchangeUser();
                    if (exUser != null)
                        temp = new Person(globalAddressList.AddressEntries[index + counter].Name, exUser.PrimarySmtpAddress.Split('@')[0]);
                    else
                        temp = new Person(globalAddressList.AddressEntries[index + counter].Name, null);

                    while (string.Compare(person.nameLastHalf[0], temp.nameLastHalf[0]) == 0)
                    {
                        if (this.Matches(person, temp))
                            matches.Add(temp);

                        Debug.WriteLine(counter);

                        if (i == 0)
                            counter++;
                        else
                            counter--;

                        exUser = globalAddressList.AddressEntries[index + counter].GetExchangeUser();

                        if (exUser != null)
                            temp = new Person(globalAddressList.AddressEntries[index + counter].Name, exUser.PrimarySmtpAddress.Split('@')[0]);
                        else
                            temp = new Person(globalAddressList.AddressEntries[index + counter].Name, null);
                    }

                    counter = 0;
                }

            Console.WriteLine("Found the following Ids for " + person.name + ":");

            foreach (Person match in matches)
                Console.WriteLine((++counter).ToString() + ". " + match.id + " - " + match.name);

            Console.WriteLine("Type the number of the correct Id or type 0 to define an Id:");

            int selection;
            int.TryParse(Console.ReadLine(), out selection);

            if (selection == 0)
            {
                Console.WriteLine("Define an Id:");
                return Console.ReadLine();
            }

            id = matches[selection - 1].id;

            return id;
        }

        private int FindLastNameIndex(int lowerBound, int upperBound)
        {
            int index = (int)Math.Floor((upperBound - lowerBound) / 2.0) + lowerBound;
            Person temp = new Person(globalAddressList.AddressEntries[index].Name, null);
            int compare = 0;

            try
            {
                compare = string.Compare(person.nameLastHalf[0], temp.nameLastHalf[0]);
            }
            catch (Exception e)
            {
                return 1;
            }

            if (compare == 0 || (upperBound - lowerBound) == 1)
                return index;
            else if (compare > 0)
                return this.FindLastNameIndex(index, upperBound);
            else
                return this.FindLastNameIndex(lowerBound, index);
        }

        private bool Matches(Person a, Person b)
        {
            int firstMatches = 0;
            int secondMatches = 0;

            for (int j = 0; j < a.nameFirstHalf.Length; j++)
                for (int k = 0; k < b.nameFirstHalf.Length; k++)
                    if (a.nameFirstHalf[j] == b.nameFirstHalf[k] && !string.IsNullOrEmpty(a.nameFirstHalf[j]) && !string.IsNullOrEmpty(b.nameFirstHalf[k]))
                        firstMatches++;

            for (int j = 0; j < a.nameLastHalf.Length; j++)
                for (int k = 0; k < b.nameLastHalf.Length; k++)
                    if (a.nameLastHalf[j] == b.nameLastHalf[k] && !string.IsNullOrEmpty(a.nameLastHalf[j]) && !string.IsNullOrEmpty(b.nameLastHalf[k]))
                        secondMatches++;

            if (firstMatches + secondMatches > 1)
                return true;
            else
                return false;
        }
    }

    public class Person
    {
        public string name;
        public string nameFirst;
        public string nameLast;
        public string[] nameFirstHalf;
        public string[] nameLastHalf;
        public string id;
        public string email;

        public Person(string n, string i)
        {
            SetId(i);
            SetName(n);
        }

        public void SetName(string n)
        {
            if (n != null)
            {
                name = this.RemoveDiacritics(n);
                SetNames();
            }
        }

        public void SetId(string i)
        {
            if (!string.IsNullOrEmpty(i))
            {
                id = i;
                email = id + "@company.com";
            }
        }

        private void SetNames()
        {
            string[] nameSplit = name.Split(',');

            if (name.Split(',').Length != 1)
            {
                nameFirst = name.Split(',')[name.Split(',').Length - 1];
                nameLast = name.Split(',')[0];
                nameLastHalf = nameLast.Split(' ');
            }
            else
            {
                nameFirst = name;
                nameLast = name;
                nameLastHalf = new string[] { name };
            }

            nameFirstHalf = nameFirst.Split(' ');
        }

        private string FindId()
        {
            IdFinder idFinder = new IdFinder(this);
            string id = idFinder.FindId();

            return id;
        }

        private string RemoveDiacritics(string s)
        {
            string stFormD = s.Normalize(NormalizationForm.FormD);
            int len = stFormD.Length;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < len; i++)
            {
                System.Globalization.UnicodeCategory uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(stFormD[i]);
                if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(stFormD[i]);
            }
            return (sb.ToString().Normalize(NormalizationForm.FormC));
        }
    }

    public class Candidate : Person
    {
        public bool onBook;
        public bool onStaff;
        public int years;
        public string[] skills;
        public int affinity;

        public Candidate(string n, string e, bool oT, bool oS, int l, string sS)
            : base(n, e)
        {
            onBook = oT;
            onStaff = oS;
            years = l;
            if (sS != null)
                skills = sS.Split(new char[] { ',', ' ', '\n', '&' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }

    public class Project
    {
        public string id;
        public string title;
        public string city;
        public string country;
        public Person contact1;
        public Person contact2;
        public List<Person> candidates = new List<Person>();
        public int yearsFrom;
        public int yearsTo;
        public string startDate;
        public string endDate;
        public int months;
        public string[] skills;
        public string[] specialities;
        public string description;

        public Project(string rI, string rT, string c, Person dS, Person rC, List<Person> cs)
        {
            id = rI;
            title = rT;
            city = c;
            contact1 = dS;
            contact2 = rC;
            candidates = cs;
        }

        public Project(string rI, string rT, string c, string cY, Person dS, Person rC, int lF, int lT, string sD, string eD, string sS, string sP, string dC)
        {
            id = rI;
            title = rT;
            city = c;
            country = cY;
            contact1 = dS;
            contact2 = rC;
            yearsFrom = lF;
            yearsTo = lT;
            startDate = sD;
            endDate = eD;
            months = this.FindMonths();
            if (sS != null)
                skills = sS.Split(new char[] { ' ', '-', ',', '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (sP != null)
                specialities = sP.Split(new string[] { " > " }, StringSplitOptions.RemoveEmptyEntries);
            description = dC;
        }

        private int FindMonths()
        {
            DateTime startDate = DateTime.ParseExact(this.startDate, "dd-MMM-yy", System.Globalization.CultureInfo.InvariantCulture);
            DateTime endDate = DateTime.ParseExact(this.endDate, "dd-MMM-yy", System.Globalization.CultureInfo.InvariantCulture);
            int months = (endDate.Month - startDate.Month) + 12 * (endDate.Year - startDate.Year);

            return months;
        }
    }

    public class ApplicationCreator
    {
        public Outlook.Application GetOutlookApplication()
        {
            Outlook.Application application = null;

            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            else
                application = new Outlook.Application();

            return application;
        }

        public Excel.Application GetExcelApplication()
        {
            Excel.Application application = null;

            if (Process.GetProcessesByName("EXCEL").Count() > 0)
                application = Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            else
                application = new Excel.Application();

            return application;
        }
    }

    public class ExcelScanner
    {
        private string matchesPath = Directory.GetCurrentDirectory() + "\\matches.xls";
        private string projectsPath = Directory.GetCurrentDirectory() + "\\projects.xls";
        private string candidatesPath = Directory.GetCurrentDirectory() + "\\candidates.xls";
        Excel.Application excelApp;

        private Excel.Range GetUsedRange(string path)
        {
            ApplicationCreator applicationCreator = new ApplicationCreator();
            excelApp = applicationCreator.GetExcelApplication();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(path);
            Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Excel.Range excelRange = excelWorksheet.UsedRange;

            return excelRange;
        }

        public List<Project> RetrieveMatches()
        {
            List<Project> matches = new List<Project>();

            Excel.Range excelRange = this.GetUsedRange(this.matchesPath);

            int rows = excelRange.Rows.Count;

            for (int i = 2; i <= rows; i++)
            {
                int sent = (int)(excelRange.Cells[i, 7] as Excel.Range).Value2;

                double dId = (excelRange.Cells[i, 1] as Excel.Range).Value2;
                string id = dId.ToString();
                string title = (string)(excelRange.Cells[i, 2] as Excel.Range).Value2;
                string city = (string)(excelRange.Cells[i, 3] as Excel.Range).Value2;

                string contact1Name = (string)(excelRange.Cells[i, 4] as Excel.Range).Value2;
                Person contact1 = new Person(contact1Name, null);

                string contact2Name = (string)(excelRange.Cells[i, 5] as Excel.Range).Value2;
                Person contact2 = new Person(contact2Name, null);

                string sCandidatesIds = (string)(excelRange.Cells[i, 6] as Excel.Range).Value2;
                string[] candidatesIds = sCandidatesIds.Split(' ');
                List<Person> candidates = new List<Person>();
                for (int j = 0; j < candidatesIds.Length; j++)
                    candidates.Add(new Person(null, candidatesIds[j]));

                matches.Add(new Project(id, title, city, contact1, contact2, candidates));
            }

            excelApp.Workbooks.Close();

            return matches;
        }

        public List<Project> RetrieveProjects()
        {
            List<Project> projects = new List<Project>();

            Excel.Range excelRange = this.GetUsedRange(projectsPath);

            int rows = excelRange.Rows.Count;
            double percent = 0;

            double idD;
            string id;
            string title;
            string city;
            string country;

            string contact1Name;
            Person contact1;

            string contact2Name;
            Person contact2;

            int yearsFrom;
            int yearsTo;

            string startDate;
            string endDate;

            string skillString;
            string specialityString;

            string description;

            for (int i = 2; i <= rows; i++)
            {
                idD = (excelRange.Cells[i, 28] as Excel.Range).Value2;
                id = idD.ToString();
                title = (string)(excelRange.Cells[i, 31] as Excel.Range).Value2;
                city = (string)(excelRange.Cells[i, 9] as Excel.Range).Value2;
                country = (string)(excelRange.Cells[i, 8] as Excel.Range).Value2;

                contact1Name = (string)(excelRange.Cells[i, 87] as Excel.Range).Value2;
                contact1 = new Person(contact1Name, null);

                contact2Name = (string)(excelRange.Cells[i, 88] as Excel.Range).Value2;
                contact2 = new Person(contact2Name, null);

                try
                {
                    yearsFrom = (int)(excelRange.Cells[i, 61] as Excel.Range).Value2;
                }
                catch (Exception)
                {
                    yearsFrom = 99;
                }
                try
                {
                    yearsTo = (int)(excelRange.Cells[i, 62] as Excel.Range).Value2;
                }
                catch (Exception)
                {
                    yearsTo = 99;
                }

                startDate = (string)(excelRange.Cells[i, 82] as Excel.Range).Value2;
                endDate = (string)(excelRange.Cells[i, 84] as Excel.Range).Value2;

                skillString = (string)(excelRange.Cells[i, 69] as Excel.Range).Value2;
                specialityString = (string)(excelRange.Cells[i, 68] as Excel.Range).Value2;

                try
                {
                    description = (string)(excelRange.Cells[i, 32] as Excel.Range).Value2;
                }
                catch (Exception)
                {
                    description = null;
                }

                percent = Math.Floor(((i - 1.0) * 100.0 / rows));

                Console.Write("\r" + percent + " / 100");
                projects.Add(new Project(id, title, city, country, contact1, contact2, yearsFrom, yearsTo, startDate, endDate, skillString, specialityString, description));
            }

            excelApp.Workbooks.Close();

            return projects;
        }

        public List<Candidate> RetrieveCandidates()
        {
            List<Candidate> candidates = new List<Candidate>();

            Excel.Range excelRange = this.GetUsedRange(candidatesPath);

            int rows = excelRange.Rows.Count;

            for (int i = 2; i <= rows; i++)
            {
                string eid = (string)(excelRange.Cells[i, 1] as Excel.Range).Value2;
                int level = (int)(excelRange.Cells[i, 2] as Excel.Range).Value2;
                int onStaff = (int)(excelRange.Cells[i, 3] as Excel.Range).Value2;
                int onBook = (int)(excelRange.Cells[i, 4] as Excel.Range).Value2;
                string skills = (string)(excelRange.Cells[i, 5] as Excel.Range).Value2;

                candidates.Add(new Candidate(null, eid, onBook.Equals(1), onStaff.Equals(1), level, skills));
            }

            excelApp.Workbooks.Close();

            return candidates;
        }
    }

    public class TextOperator
    {
        private string sentPath = Directory.GetCurrentDirectory() + "\\sent.txt";
        private StreamWriter writer;

        public void WriteSentProject(Project project)
        {
            writer = new StreamWriter(sentPath, true);
            writer.WriteLine(project.id);
            writer.Close();
        }

        public string[] GetSentProjects()
        {
            string[] sent = System.IO.File.ReadAllLines(sentPath);

            return sent;
        }
    }

    public class Matcher
    {
        string[] skills;
        string[] specialities;

        public Matcher()
        {
            skills = File.ReadAllText("skills.txt").Split(',');
            specialities = File.ReadAllText("specialities.txt").Split(',');
        }

        public List<Project> Match()
        {
            List<Project> matches = new List<Project>();
            List<Candidate> tempCandidates;
            List<Candidate> sortedCandidates;
            int skillMatches = 0;
            ExcelScanner excelScanner = new ExcelScanner();
            List<Project> projects = excelScanner.RetrieveProjects();
            List<Candidate> candidates = excelScanner.RetrieveCandidates();
            MailSender mailSender = new MailSender();
            string shouldBuildEmail;
            int count = 0;
            int tolerance;

            foreach (Project project in projects)
            {
                ConsoleWriter.WriteLine("Matching " + project.id + " " + project.title + " (" + project.yearsFrom + "-" + project.yearsTo + ")...", 'y');
                ConsoleWriter.WriteLine("Project description :" + project.description, 'y');

                if (this.IsItTimely(project) && !this.IsItSent(project) && !this.RequiresCitizenship(project))
                {
                    tempCandidates = new List<Candidate>();

                    foreach (Candidate candidate in candidates)
                    {
                        skillMatches = 0;

                        if (this.IsItRelevant(project))
                            tolerance = 1;
                        else
                            tolerance = 3;

                        if (candidate.years >= (project.yearsFrom - 1) && candidate.years <= (project.yearsTo + 1) && candidate.onStaff && candidate.onBook && candidate.skills != null)
                            if (project.skills != null)
                            {
                                Console.WriteLine("Comparing " + candidate.id + "...");
                                skillMatches = this.CompareSkills(project.skills, candidate.skills);
                            }
                            else
                                skillMatches = 1;

                        if (skillMatches >= tolerance)
                        {
                            candidate.affinity = skillMatches;
                            tempCandidates.Add(candidate);
                        }
                    }

                    if (tempCandidates.Count != 0)
                    {
                        sortedCandidates = tempCandidates.OrderByDescending(o => o.affinity).ToList();

                        foreach (Candidate sortedCandidate in sortedCandidates)
                            if (project.candidates.Count < 4)
                            {
                                ConsoleWriter.WriteLine("Candidate " + sortedCandidate.id + " (" + sortedCandidate.years + ") added!", 'b');
                                project.candidates.Add(sortedCandidate);
                            }
                    }

                    if (project.candidates.Count != 0)
                    {
                        ConsoleWriter.WriteLine("Job added!", 'b');
                        matches.Add(project);
                        count++;

                        Console.WriteLine("Build email? (y/n)");
                        shouldBuildEmail = Console.ReadLine();

                        if (shouldBuildEmail == "y")
                            mailSender.SendMail(project);
                    }
                }
            }

            return matches;
        }

        private bool IsItRelevant(Project project)
        {
            bool isItRelevant = false;
            int skillMatches;
            int specialityMatches;
            int nameMatches;

            Console.WriteLine("Comparing skills...");
            skillMatches = this.CompareSkills(project.skills, this.skills);
            Console.WriteLine("Comparing specialities...");
            specialityMatches = this.CompareSkills(project.specialities, this.specialities);
            Console.WriteLine("Comparing name...");
            nameMatches = this.CompareSkills(project.title.Split(new char[] { ' ', '-', ',', '/', '<', '>' }, StringSplitOptions.RemoveEmptyEntries), this.skills);

            if (skillMatches > 1 || specialityMatches >= 1 || (nameMatches >= 1 && skillMatches >= 1))
            {
                ConsoleWriter.WriteLine("Project IS relevant.", 'b');
                isItRelevant = true;
            }
            else
                ConsoleWriter.WriteLine("Project is NOT relevant.", 'r');

            return isItRelevant;
        }

        private bool IsItSent(Project project)
        {
            bool isItSent = false;
            TextOperator textOperator = new TextOperator();
            string[] sent = textOperator.GetSentProjects();

            foreach (string sentID in sent)
                if (sentID == project.id)
                {
                    isItSent = true;
                    ConsoleWriter.WriteLine("Project has been sent already!", 'r');
                }

            return isItSent;
        }

        private bool IsItTimely(Project project)
        {
            bool isItTimely = true;
            if (project.country == "USA" && project.months > 6)
            {
                isItTimely = false;
                ConsoleWriter.WriteLine("Project is not Timely!", 'r');
            }

            return isItTimely;
        }

        private bool RequiresCitizenship(Project project)
        {
            bool requiresCitizenship = false;
            string[] keyWords = { "citizen", "citizenship", "federal", "citizens" };
            string[] descriptionWords;

            if (project.description != null)
            {
                descriptionWords = project.description.Split(new char[] { ',', ' ', '.', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                if (this.CompareSkills(keyWords, descriptionWords) > 0)
                {
                    requiresCitizenship = true;
                    ConsoleWriter.WriteLine("Project requires citizenship!", 'r');
                }
            }

            return requiresCitizenship;
        }

        private int CompareSkills(string[] target, string[] compared)
        {
            List<string> found = new List<string>();
            List<string> distinctFound;

            if (target != null)
                foreach (string targetItem in target)
                    foreach (string comparedItem in compared)
                        if (targetItem.ToLower() == comparedItem.ToLower() && targetItem != "" && comparedItem != "" && targetItem.ToLower() != "business" && targetItem.ToLower() != "and")
                            found.Add(targetItem.ToLower());

            distinctFound = found.Distinct().ToList();

            foreach (string foundItem in distinctFound)
                Console.WriteLine("Match: " + foundItem);

            return distinctFound.Count;
        }
    }

    public static class ConsoleWriter
    {
        public static void WriteLine(string s, char color)
        {
            if (color == 'b')
                Console.ForegroundColor = ConsoleColor.Blue;
            else if (color == 'y')
                Console.ForegroundColor = ConsoleColor.Yellow;
            else if (color == 'r')
                Console.ForegroundColor = ConsoleColor.Red;

            Console.WriteLine(s);
            Console.ResetColor();
        }
    }
}
