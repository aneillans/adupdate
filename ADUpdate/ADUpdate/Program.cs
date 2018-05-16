using CommandLine;
using Dynamitey;
using FileHelpers;
using FileHelpers.Dynamic;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;

namespace ADUpdate
{
    class Program
    {
        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(opts => RunOptionsAndReturnExitCode(opts))
                .WithNotParsed<Options>((errs) => HandleParseError(errs));
        }

        private static void RunOptionsAndReturnExitCode(Options opts)
        {
            if (!File.Exists(opts.InputFile) || !File.Exists(opts.MappingFile))
            {
                Console.WriteLine("Input and mapping definition files are required for processing");
                return;
            }

            if (string.IsNullOrEmpty(opts.Domain))
            {
                opts.Domain = Environment.UserDomainName;
                Console.WriteLine(" ==> Domain has been set to: " + opts.Domain);
            }

            Dictionary<string, string> fieldMappings = new Dictionary<string, string>();

            using (StreamReader sr = new StreamReader(opts.MappingFile))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    var split = line.Split('=');
                    if (split.Length == 2 && !string.IsNullOrEmpty(split[0]) && !string.IsNullOrEmpty(split[1]))
                    {
                        fieldMappings.Add(split[0], split[1]);
                    }
                }
            }

            var detector = new FileHelpers.Detection.SmartFormatDetector();
            detector.MaxSampleLines = 10;
            var formats = detector.DetectFileFormat(opts.InputFile);
            bool keyFound = false;
            DelimitedClassBuilder dcb;
            DelimitedFileEngine fileEngine = null;

            foreach (var format in formats)
            {
                Console.WriteLine("File Format Detected, confidence:" + format.Confidence + "%");
                var delimited = format.ClassBuilderAsDelimited;
                Console.WriteLine("    Delimiter:" + delimited.Delimiter);
                dcb = new DelimitedClassBuilder("InputFile", delimited.Delimiter);
                Console.WriteLine("    Fields:");

                foreach (var field in delimited.Fields)
                {
                    dcb.AddField(field);
                    var mappedField = "<Ignored>";
                    if (fieldMappings.ContainsKey(field.FieldName))
                    {
                        mappedField = fieldMappings[field.FieldName];

                    }
                    var keyField = string.Empty;
                    if (field.FieldName == opts.Key)
                    {
                        keyField = "  << UNIQUE KEY >>";
                        keyFound = true;
                    }

                    Console.WriteLine("        " + field.FieldName + ": " + field.FieldType + " => " + mappedField + keyField);
                }

                fileEngine = new DelimitedFileEngine(dcb.CreateRecordClass());
            }

            if (!keyFound)
            {
                Console.WriteLine("The specified unique key field, " + opts.Key + ", has not been found in the input file - aborting.");
                return;
            }

            if (fileEngine == null)
            {
                return;
            }

            DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://" + opts.Domain, null, null, AuthenticationTypes.Secure);
            DirectorySearcher search = new DirectorySearcher(ldapConnection);
            foreach (var mapping in fieldMappings)
            {
                search.PropertiesToLoad.Add(mapping.Value);
            }
            dynamic[] sourceFile = fileEngine.ReadFile(opts.InputFile);
            Console.WriteLine();
            Console.WriteLine("Input File loaded; " + (sourceFile.Length - 1) + " records");
            bool firstLineSkipped = false;
            int count = 1;
            foreach (dynamic rec in sourceFile)
            {
                if (firstLineSkipped)
                {
                    string key = Dynamic.InvokeGet(rec, opts.Key);
                    search.Filter = "(employeeID=" + key + ")";
                    SearchResultCollection result = search.FindAll();
                    if (result.Count == 1)
                    {
                        var user = result[0].GetDirectoryEntry();
                        Console.WriteLine(count + "/" + (sourceFile.Length - 1) + ": Checking user " + user.Name);

                        foreach (var mapping in fieldMappings)
                        {
                            if (mapping.Value == opts.Key)
                            {
                                continue;
                            }
                            bool updateValue = false;
                            string oldValue = string.Empty;
                            string newValue = string.Empty;
                            newValue = Dynamic.InvokeGet(rec, mapping.Key);
                            if (user.Properties[mapping.Value].Value != null)
                            {
                                if (mapping.Value.ToLower() == "manager")
                                {
                                    if (!string.IsNullOrEmpty(Dynamic.InvokeGet(rec, mapping.Key)))
                                    {
                                        // Manager has to be treated differently, as its a DN reference to another AD object.
                                        string[] man = Dynamic.InvokeGet(rec, mapping.Key).Split(' ');

                                        // Lookup what the manager DN SHOULD be
                                        DirectorySearcher managerSearch = new DirectorySearcher(ldapConnection);
                                        managerSearch.PropertiesToLoad.Add("distinguishedName");
                                        managerSearch.Filter = "(&(givenName=" + man[0] + ")(sn=" + man[1] + "))";
                                        var manager = managerSearch.FindOne();
                                        if (manager != null)
                                        {
                                            newValue = manager.GetDirectoryEntry().Properties["distinguishedName"].Value.ToString();

                                            if (user.Properties[mapping.Value].Value.ToString() != newValue)
                                            {
                                                updateValue = true;
                                                oldValue = user.Properties[mapping.Value].Value.ToString();
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (user.Properties[mapping.Value].Value.ToString() != Dynamic.InvokeGet(rec, mapping.Key))
                                    {
                                        updateValue = true;
                                        oldValue = user.Properties[mapping.Value].Value.ToString();
                                    }
                                }
                            }
                            else
                            {
                                updateValue = true;
                            }

                            if (updateValue)
                            {
                                if (mapping.Value.ToLower() == "manager")
                                {
                                    if (!string.IsNullOrEmpty(Dynamic.InvokeGet(rec, mapping.Key)))
                                    {
                                        // Manager has to be treated differently, as its a DN reference to another AD object.
                                        string[] man = Dynamic.InvokeGet(rec, mapping.Key).Split(' ');

                                        // Lookup what the manager DN SHOULD be
                                        DirectorySearcher managerSearch = new DirectorySearcher(ldapConnection);
                                        managerSearch.PropertiesToLoad.Add("distinguishedName");
                                        managerSearch.Filter = "(&(givenName=" + man[0] + ")(sn=" + man[1] + "))";
                                        var manager = managerSearch.FindOne();
                                        if (manager != null)
                                        {
                                            newValue = manager.GetDirectoryEntry().Properties["distinguishedName"].Value.ToString();
                                        }
                                    }
                                }

                                Console.WriteLine(" >> Value to be updated (" + mapping.Value + ") - " + oldValue + " != " + newValue);
                                if (!opts.WhatIf)
                                {
                                    user.Properties[mapping.Value].Clear();
                                    user.Properties[mapping.Value].Add(newValue);
                                }
                            }
                        }

                        if (!opts.WhatIf)
                        {
                            user.CommitChanges();
                            Console.WriteLine("Updated user " + user.Name);
                        }
                    }
                    else
                    {
                        if (result.Count > 1)
                        {
                            Console.WriteLine(count + "/" + (sourceFile.Length - 1) + ": -> Multiple entries returned; " + key);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine(count + "/" + (sourceFile.Length - 1) + ": -> No record found for " + key);
                            if (opts.LooseMatching)
                            {
                                string userFullName = string.Empty;
                                string firstName = string.Empty;
                                string lastName = string.Empty;
                                foreach (var mapping in fieldMappings)
                                {
                                    if (mapping.Value == "givenName")
                                    {
                                        firstName = Dynamic.InvokeGet(rec, mapping.Key);
                                    }
                                    else if (mapping.Value == "sn")
                                    {
                                        lastName = Dynamic.InvokeGet(rec, mapping.Key);
                                    }
                                }
                                if (string.IsNullOrEmpty(firstName) || string.IsNullOrEmpty(lastName))
                                {
                                    Console.WriteLine("No firstname and lastname values found mapped, can not loosely match - aborting run");
                                    return;
                                }
                                userFullName = firstName + " " + lastName;
                                Console.WriteLine(" -> Trying search for users full name " + userFullName);
                                search.Filter = "(cn=" + userFullName + ")";
                                result = search.FindAll();
                                if (result.Count == 1)
                                {
                                    var user = result[0].GetDirectoryEntry();
                                    Console.WriteLine("Checking user " + user.Name);

                                    foreach (var mapping in fieldMappings)
                                    {
                                        bool updateValue = false;
                                        if (user.Properties[mapping.Value].Value != null)
                                        {
                                            if (user.Properties[mapping.Value].Value.ToString() != Dynamic.InvokeGet(rec, mapping.Key))
                                            {
                                                updateValue = true;
                                            }
                                        }
                                        else
                                        {
                                            updateValue = true;
                                        }

                                        if (updateValue)
                                        {
                                            Console.WriteLine("Value to be updated (" + mapping.Value + ") ");
                                            if (!opts.WhatIf)
                                            {
                                                user.Properties[mapping.Value].Clear();
                                                user.Properties[mapping.Value].Add(Dynamic.InvokeGet(rec, mapping.Key));
                                            }
                                        }
                                    }

                                    if (!opts.WhatIf)
                                    {
                                        user.CommitChanges();
                                        Console.WriteLine("Updated user " + user.Name);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(" E -> Too many or no hits to for a loose match");
                                }
                            }
                        }
                    }
                    count++;
                }
                firstLineSkipped = true;
            }
        }


        private static object HandleParseError(IEnumerable<Error> errs)
        {
            foreach (Error e in errs)
            {
                Console.WriteLine("Error: " + e);
            }
            return null;
        }

        class Options
        {
            [Option('i', "input", Required = true,
              HelpText = "CSV Input file to be processed.")]
            public string InputFile { get; set; }

            [Option('m', "mappings", Required = true,
                HelpText = "Mapping Definition file.")]
            public string MappingFile { get; set; }

            [Option(
              HelpText = "Operates in a WhatIf mode - no changes will be applied.")]
            public bool WhatIf { get; set; }

            [Value(0, MetaName = "key", Required = true,
              HelpText = "Field to use as primary key")]
            public string Key { get; set; }

            [Option(
                HelpText = "Allow loose matching on full name if no key found")]
            public bool LooseMatching { get; set; }

            [Option('d', "domain",
                HelpText = "Domain to update against"
                )]
            public string Domain { get; set; }
        }
    }
}
