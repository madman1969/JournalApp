using JournalApp.DataModel;
using System;
using System.Collections.Generic;

namespace JournalApp.CmdLine
{
    internal class Program
    {
        static void Main(string[] args)
        {
            const string jsonFile = "journal.json";
            List<JournalEntry> entryList = new List<JournalEntry>();

            // entryList = JournalApp.Parser.Parser.RetrieveJournalEntriesFromOutlook(jsonFile);
            entryList = Parser.Parser.RetrieveJournalEntriesFromJSONFile(jsonFile);

            Console.WriteLine();
        }
    }
}
