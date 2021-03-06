﻿namespace Linx
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;

    using ColoredConsole;

    public class Program
    {
        private const string ExclusionsKey = "linx_exclusions";

        public static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            string[] inputs = null;
            var outputs = new List<string>();
            var format = OutputFormat.md;

            if (args?.Length > 0)
            {
                inputs = ProcessInputs(ref format, args);
            }
            else
            {
                ColorConsole.Write("Provide the files to parse links and the format to save", " (e.g. c:\\deckWithLinks.pptx;c:\\docWithLinks.docx;c:\\folderWithDocs --md)".DarkGray(), ": ".Green());
                var param = Console.ReadLine().Split("--", StringSplitOptions.RemoveEmptyEntries);
                inputs = ProcessInputs(ref format, param);
            }

            InputBase.Exclusions = (Environment.GetEnvironmentVariable(ExclusionsKey, EnvironmentVariableTarget.User) ?? Environment.GetEnvironmentVariable(ExclusionsKey, EnvironmentVariableTarget.Machine))?.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)?.Select(x => x.Trim()).ToList();
            foreach (var input in inputs)
            {
                if (File.Exists(input))
                {
                    ProcessFile(format, outputs, input.Trim());
                }
                else
                {
                    if (Directory.Exists(input))
                    {
                        var files = new[] { "*.pptx", "*.docx" }.SelectMany(x => Directory.EnumerateFiles(input.Trim(), x, SearchOption.AllDirectories));
                        foreach (var file in files)
                        {
                            ProcessFile(format, outputs, file.Trim());
                        }
                    }
                }
            }

            ColorConsole.WriteLine("outputs", ": ".Green(), outputs.Count.ToString().DarkGray());
            if (outputs?.Count > 0)
            {
                if (outputs.Count > 1)
                {
                    ColorConsole.Write("\n", "> ".Green(), "Merge outputs? ", "(y/n) ".Green());
                    var yesno = Console.ReadKey();
                    ColorConsole.WriteLine();
                    if (yesno.Key == ConsoleKey.Y)
                    {
                        OutputBase.GetInstance(format).Merge(outputs);
                    }
                }

                outputs.ForEach(o => ColorConsole.WriteLine(o.DarkGray()));
            }

            ColorConsole.Write("\nPress any key to continue", "...".Green());
            Console.ReadLine();
        }

        private static string[] ProcessInputs(ref OutputFormat format, string[] param)
        {
            var inputs = param.Where(p => !string.IsNullOrWhiteSpace(p)).FirstOrDefault().TrimStart('-').Trim().Split(new[] { ';', ',', '|' }, StringSplitOptions.RemoveEmptyEntries);
            if (param.Length > 1)
            {
                Enum.TryParse(param.LastOrDefault().TrimStart('-'), out format);
            }

            return inputs;
        }

        private static void ProcessFile(OutputFormat format, List<string> outputs, string file)
        {
            ColorConsole.WriteLine("input", ": ".Green(), file.DarkGray());
            ColorConsole.Write("> ");
            try
            {
                var results = InputBase.GetInstance(file).ExtractLinks(file);
                var outputFile = Path.Combine(Path.GetDirectoryName(file), $"{$"{nameof(Linx)}_{Path.GetFileName(file)}."}{format}");
                new ConsoleOut().Save(results, outputFile);
                if (OutputBase.GetInstance(format).Save(results, outputFile))
                {
                    outputs.Add(outputFile);
                }
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed());
            }
        }
    }
}
