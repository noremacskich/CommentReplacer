﻿using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommentReplacer
{
    class Program
    {
        static void Main(string[] args)
        {

            if(args.Length!= 3)
            {
                Console.WriteLine("You need at least three parameters: .\\CommentReplacer \"Path to file\" \"word or words to replace\" \"Comment to apply\"");
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
    }
}