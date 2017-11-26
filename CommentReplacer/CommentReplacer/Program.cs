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

            if (args.Length != 3)
            {
                Console.WriteLine("You need the following three parameters: .\\CommentReplacer \"Path to file\" \"word or words to replace\" \"Comment to apply\"");
                Environment.Exit(1);
            }

            string filePath = args[0];
            string wordToReplace = args[1];
            string commentToAdd = args[2];

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

            Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;


            Document thisDocument = wordApp.Documents.Open(@filePath);
            // https://msdn.microsoft.com/en-us/library/e7d13z59.aspx
            Range rng = thisDocument.Content;

            rng.Find.ClearFormatting();
            rng.Find.Forward = true;
            rng.Find.Text = wordToReplace;
            rng.Find.MatchWholeWord = true;

            rng.Find.Execute();

            while (rng.Find.Found)
            {
                // Make sure that we are not putting duplicate comments on a particular range
                if (!AlreadyHasComment(rng, commentToAdd))
                {
                    thisDocument.Comments.Add(rng, commentToAdd);
                }
                // Move on to the next finding
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
                using (Stream stream = new FileStream(pathToFile, FileMode.Open))
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
        /// <summary>
        /// Will return true if the comment already exists on the range, false if it doesn't.
        /// </summary>
        /// <param name="thisRange">The range to check.</param>
        /// <param name="commentToAdd">The comment you are checking for.</param>
        /// <returns></returns>
        private static bool AlreadyHasComment(Range thisRange, string commentToAdd)
        {
            foreach (Comment thisComment in thisRange.Comments)
            {
                if (thisComment.Range.Text == commentToAdd)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
