﻿using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace JournalApp
{
  internal class Program
  {
    static void Main(string[] args)
    {
      const string jsonFile = "journal.json";
      List<JournalEntry> entryList = new List<JournalEntry>();

      try
      {
        JournalEntry tmp = new JournalEntry();

        Microsoft.Office.Interop.Outlook.JournalItem current;

        // Create the Outlook application.
        // in-line initialization
        Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();

        // Get the MAPI namespace.
        Microsoft.Office.Interop.Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

        // Log on by using the default profile or existing session (no dialog box).
        oNS.Logon(Missing.Value, Missing.Value, false, true);

        // Get list of journal entries in descending creation date order ...
        var journalFolder = oApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderJournal);
        var journalItems = journalFolder.Items;
        journalItems.Sort("[CreationTime]", true);

        Console.WriteLine($"Found {journalFolder.Items.Count} journal entries");

        foreach (var entry in journalItems)
        {
          current = (Microsoft.Office.Interop.Outlook.JournalItem)entry;
          Console.WriteLine($"\tJournal Subject: [{current.Subject}] [Journey Type: {current.Type}] [{current.Attachments.Count} Attachments] [Creation Date: {current.CreationTime.ToString()}]");

          if (current.Attachments.Count > 0)
          {
            var attachmentsList = (Microsoft.Office.Interop.Outlook.Attachments)current.Attachments;

            foreach (Attachment attachment in attachmentsList)
            {
              Console.WriteLine($"\t[{attachment.Type.ToString()}] [{attachment.DisplayName}]");
            }
          }

          entryList.Add(new JournalEntry
          {
            Subject = current.Subject,
            EntryType = current.Type,
            StartTime = current.CreationTime,
            Body = current.Body
          });
        }

        //Log off.
        oNS.Logoff();

        Console.WriteLine("Logged out of Outlook");

        // Write journal list out as JSON ...
        string json = JsonConvert.SerializeObject(entryList, Newtonsoft.Json.Formatting.Indented);
        var output = Path.Combine(GetExecutingDirectoryName(), jsonFile);

        Console.WriteLine($"JSON File Path: [{output}]");

        if (File.Exists(output))
          File.Delete(output);

        File.WriteAllText(output, json);

      }
      //Error handler.
      catch (System.Exception e)
      {
        Console.WriteLine("{0} Exception caught: ", e);
      }
    }

    public static string GetExecutingDirectoryName()
    {
      var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
      return new FileInfo(location.AbsolutePath).Directory.FullName;
    }
  }
}