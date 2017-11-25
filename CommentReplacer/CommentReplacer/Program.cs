using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CommentReplacer
{
    class Program
    {
        static void Main(string[] args)
        {

            switch (args[0])
            {
                case "-v":
                case "v":
                    Console.WriteLine("Comment Replacer V 1.0");
                    Console.WriteLine("Author: NoremacSkich");
                    Environment.Exit(0);
                    break;
                case "help":
                case "-h":
                case "h":
                    Console.WriteLine("The purpose of this program is to add a comment to all instances of the word specified.");
                    Console.WriteLine("You need the following three parameters: .\\CommentReplacer \"Path to file\" \"word or words to replace\" \"Comment to apply\"");
                    Environment.Exit(0);
                    break;

            }

            if(args.Length!= 3)
            {
                Console.WriteLine("You need the following three parameters: .\\CommentReplacer \"Path to file\" \"word or words to replace\" \"Comment to apply\"");
                Environment.Exit(1);
            }



            string filePath = args[0];
            string wordToReplace = args[1];
            string commentToAdd = args[2];

            Console.WriteLine(filePath);
            Console.WriteLine(wordToReplace);
            Console.WriteLine(commentToAdd);


            Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            // Need to make sure that file exists
            // Need to make sure that it's not already open 

            if (!File.Exists(filePath))
            {
                Console.WriteLine("This file doesn't exist, exiting application.");
                Environment.Exit(1);
            }

            if (IsAlreadyOpen(filePath))
            {
                Console.WriteLine("This file is already open, exiting application");
                Environment.Exit(1);
            }

            Document thisDocument = wordApp.Documents.Open(@filePath);
            // https://msdn.microsoft.com/en-us/library/e7d13z59.aspx
            Range rng = thisDocument.Content;

            rng.Find.ClearFormatting();
            rng.Find.Forward = true;
            rng.Find.Text = wordToReplace;

            rng.Find.Execute();

            while (rng.Find.Found)
            {
                thisDocument.Comments.Add(rng, commentToAdd);
                rng.Find.Execute();
            }

            thisDocument.Save();
            thisDocument.Close();
            wordApp.Quit();
        }
        //https://stackoverflow.com/a/876513/3271665
        private static bool IsAlreadyOpen(string pathToFile)
        {
            try
            {
                using (Stream stream = new FileStream("MyFilename.txt", FileMode.Open))
                {
                    // File/Stream manipulating code here
                }
                return false;
            }
            catch
            {
                //check here why it failed and ask user to retry if the file is in use.
                return true;
            }
        }
    }
}
