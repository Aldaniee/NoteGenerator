using System;
using System.Diagnostics;
using System.Globalization;
using Xceed.Words.NET;

namespace NoteGenerator {
    class Program {

        static void Main(string[] args) {
            string name = "Aidan Lee";
            string[] classes = {"Intermediate Spanish 1", "Digital Circuits & Systems", "High-Performance Computing"};
            DateTime dt = DateTime.Now;
            string date = dt.ToString("yyyy-MM-dd");
            string displayDate = dt.ToString("M / d / yyyy");

            Console.WriteLine($"Creating a new note for {name}...");
            Console.Write("Title: ");
            string title = Console.ReadLine();
            TextInfo ti = new CultureInfo("en-US", false).TextInfo;
            Console.WriteLine("\nClasses: ");
            for(int i = 0; i < classes.Length; i++)
                Console.WriteLine($"\t {i}: {classes[i]}");
            Console.Write("Class: ");
            string className = classes[Convert.ToInt32(Console.ReadLine())];
            string fileName = $@"C:\Users\Aidan Lee\Google Drive\Trinity\Junior\{className}\{date} Notes.docx";
            string body = "";
            var doc = DocX.Create(fileName);

            //Formatting Title  
            Formatting headerFormat = new Formatting();
            //Specify font properties
            headerFormat.FontFamily = new Font("Cambria");
            headerFormat.Size = 12;
            headerFormat.FontColor = System.Drawing.Color.Black;

            //Insert text  
            doc.InsertParagraph(name, false, headerFormat);
            doc.InsertParagraph(displayDate, false, headerFormat);
            doc.InsertParagraph(className, false, headerFormat);

            //Specify font properties
            Formatting titleFormat = new Formatting();
            titleFormat.FontFamily = new Font("Cambria");
            titleFormat.Size = 12;
            titleFormat.FontColor = System.Drawing.Color.Black;
            titleFormat.UnderlineColor = System.Drawing.Color.Black;
            Paragraph par = doc.InsertParagraph(ti.ToTitleCase(title), false, titleFormat);
            par.Alignment = Alignment.center;

            doc.InsertParagraph(body, false, headerFormat);
            doc.Save();

            Process.Start(fileName);

        }
    }
}