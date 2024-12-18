using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

if (args.Length == 0)
{
    Console.WriteLine("Please provide the file path as an argument.");
    return;
}

string filepath = args[0];

string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
string logFilePath = $"log_{timestamp}.txt";
var logStream = new FileStream(logFilePath, FileMode.Create);
var logWriter = new StreamWriter(logStream) { AutoFlush = true };
Console.SetOut(logWriter);

ValidateWordDocument(filepath);

static void ValidateWordDocument(string filepath)
{
    using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
    {
        try
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;
            foreach (ValidationErrorInfo error in validator.Validate(wordprocessingDocument))
            {
                count++;
                Console.WriteLine("Error " + count);
                Console.WriteLine("Description: " + error.Description);
                Console.WriteLine("ErrorType: " + error.ErrorType);
                Console.WriteLine("Node: " + error.Node);
                if (error.Path is not null)
                {
                    Console.WriteLine("Path: " + error.Path.XPath);
                }
                if (error.Part is not null)
                {
                    Console.WriteLine("Part: " + error.Part.Uri);
                }
                Console.WriteLine("-------------------------------------------");
            }

            Console.WriteLine("count={0}", count);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}
