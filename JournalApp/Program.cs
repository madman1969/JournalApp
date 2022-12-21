﻿using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using ShellProgressBar;
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
      int totalTicks;

      try
      {
        JournalEntry tmp = new JournalEntry();
        JournalItem current;

        // Create the Outlook application in-line initialization
        Application oApp = new Application();

        // Get the MAPI namespace.
        NameSpace oNS = oApp.GetNamespace("mapi");

        // Log on by using the default profile or existing session (no dialog box).
        oNS.Logon(Missing.Value, Missing.Value, false, true);

        // Get list of journal entries in descending creation date order ...
        var journalFolder = oApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderJournal);
        var journalItems = journalFolder.Items;
        journalItems.Sort("[CreationTime]", true);

        totalTicks = journalFolder.Items.Count;
        //Console.WriteLine($"Found {totalTicks} journal entries");
        //Console.WriteLine();

        var options = new ProgressBarOptions
        {
          ForegroundColor = ConsoleColor.Green,
          ForegroundColorDone = ConsoleColor.DarkGreen,
          BackgroundColor = ConsoleColor.DarkGray,
          BackgroundCharacter = '\u2593',
          // ProgressCharacter = '-',
        };

        using (var pbar = new ProgressBar(totalTicks, $"Processing {totalTicks} journal entries", options))
        {
          // Generate list of JournalEntry instances ...
          foreach (var entry in journalItems)
          {
            current = (JournalItem)entry;
            // Console.WriteLine($"Journal Subject: [{current.Subject}] [Journey Type: {current.Type}] [{current.Attachments.Count} Attachments] [Creation Date: {current.CreationTime.ToString()}]");

            if (current.Attachments.Count > 0)
            {
              var attachmentsList = (Attachments)current.Attachments;

              //foreach (Attachment attachment in attachmentsList)
              //{
              //  Console.WriteLine($"\t[{attachment.Type.ToString()}] [{attachment.DisplayName}]");
              //}
            }

            entryList.Add(new JournalEntry
            {
              ConversationID = current.ConversationID,
              Subject = current.Subject,
              EntryType = current.Type,
              StartTime = current.CreationTime,
              Body = current.Body
            });

            pbar.Tick();
          }
        }

        //Log off.
        oNS.Logoff();

        Console.WriteLine("Logged out of Outlook");

        // Write journal list out as JSON ...
        WriteJSONFile(jsonFile, entryList);

      }
      //Error handler.
      catch (System.Exception e)
      {
        Console.WriteLine("{0} Exception caught: ", e);
      }
    }

    /// <summary>
    /// Writes the list of journal entries to file as JSON
    /// </summary>
    /// <param name="jsonFile"></param>
    /// <param name="entryList"></param>
    private static void WriteJSONFile(string jsonFile, List<JournalEntry> entryList)
    {
      string json = JsonConvert.SerializeObject(entryList, Formatting.Indented);
      var output = Path.Combine(GetExecutingDirectoryName(), jsonFile);

      if (File.Exists(output))
        File.Delete(output);

      File.WriteAllText(output, json);

      var fileSize = GetFileSize(output) / 1024;
      Console.WriteLine($"JSON File Path: [{output}], Size: [{fileSize}KB]");
    }

    /// <summary>
    /// Retrieve the size in bytes of the specified file
    /// </summary>
    /// <param name="FilePath"></param>
    /// <returns></returns>
    static long GetFileSize(string FilePath)
    {
      if (File.Exists(FilePath))
      {
        return new FileInfo(FilePath).Length;
      }
      return 0;
    }

    /// <summary>
    /// Retrieve directory path of the executing application
    /// </summary>
    /// <returns></returns>
    public static string GetExecutingDirectoryName()
    {
      var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
      return new FileInfo(location.AbsolutePath).Directory.FullName;
    }
  }
}
