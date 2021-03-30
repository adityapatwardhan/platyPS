using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.PowerShell.PlatyPS.Model;

namespace Microsoft.PowerShell.PlatyPS.MarkdownWriter
{
    internal class CommandHelpMarkdownWriter
    {
        private readonly string _filePath;
        private StringBuilder sb = null;
        private readonly Encoding _encoding;
        private readonly CommandHelp _help;

        public CommandHelpMarkdownWriter(string path, CommandHelp commandHelp, Encoding encoding)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException(nameof(path));
            }
            else
            {
                _filePath = path;
                _help = commandHelp;
                _encoding = encoding;
            }
        }

        internal FileInfo Write()
        {
            sb ??= new StringBuilder();

            WriteMetadataHeader();
            sb.AppendLine();

            WriteTitle();
            sb.AppendLine();

            WriteSynopsis();
            sb.AppendLine();

            // this adds an empty line after all parameters
            WriteSyntax();

            WriteDescription();
            sb.AppendLine();

            WriteExamples();
            sb.AppendLine();

            WriteParameters();

            WriteInputsOutputs(_help.Inputs, Constants.InputsMdHeader);

            WriteInputsOutputs(_help.Outputs, Constants.OutputsMdHeader);

            WriteNotes();

            WriteRelatedLinks();

            using (StreamWriter mdFileWriter = new(_filePath, append: false, _encoding))
            {
                mdFileWriter.Write(sb.ToString());

                return new FileInfo(_filePath);
            }
        }

        private void WriteMetadataHeader()
        {
            sb.AppendLine(Constants.YmlHeader);
            sb.AppendLine($"external help file: {_help.ModuleName}-help.xml");
            sb.AppendLine($"Module Name: {_help.ModuleName}");
            sb.AppendLine("online version:");
            sb.AppendLine(Constants.SchemaVersionYml);
            sb.AppendLine(Constants.YmlHeader);
        }

        private void WriteTitle()
        {
            sb.AppendLine($"# {_help.Title}");
        }

        private void WriteSynopsis()
        {
            sb.AppendLine(Constants.SynopsisMdHeader);
            sb.AppendLine();
            sb.AppendLine(_help.Synopsis);
        }

        private void WriteSyntax()
        {
            sb.AppendLine(Constants.SyntaxMdHeader);
            sb.AppendLine();

            foreach(SyntaxItem item in _help.Syntax)
            {
                sb.AppendLine(item.ToSyntaxString());
            }
        }

        private void WriteDescription()
        {
            sb.AppendLine(Constants.DescriptionMdHeader);
            sb.AppendLine();
            sb.AppendLine(_help.Description);
        }

        private void WriteExamples()
        {
            sb.AppendLine(Constants.ExamplesMdHeader);
            sb.AppendLine();

            int totalExamples = _help.Examples.Count;

            for(int i = 0; i < totalExamples; i++)
            {
                sb.Append(_help.Examples[i].ToExampleItemString(i + 1));
                sb.AppendLine();
            }
        }

        private void WriteParameters()
        {
            sb.AppendLine(Constants.ParametersMdHeader);
            sb.AppendLine();

            // Sort the parameter by name before writing
            _help.Parameters.Sort((u1, u2) => u1.Name.CompareTo(u2.Name));

            foreach(var param in _help.Parameters)
            {
                string paramString = param.ToParameterString();

                if (!string.IsNullOrEmpty(paramString))
                {
                    sb.AppendLine(paramString);
                    sb.AppendLine();
                }
            }

            sb.AppendLine(Constants.CommonParameters);
        }

        private void WriteInputsOutputs(List<InputOutput> inputsoutputs, string header)
        {
            sb.AppendLine(header);
            sb.AppendLine();

            if (inputsoutputs == null)
            {
                return;
            }

            foreach (var item in inputsoutputs)
            {
                sb.Append(item.ToInputOutputString());
            }
        }

        private void WriteNotes()
        {
            sb.AppendLine(Constants.NotesMdHeader);
            sb.AppendLine();
            sb.AppendLine(_help.Notes);
            sb.AppendLine();
        }

        private void WriteRelatedLinks()
        {
            sb.AppendLine(Constants.RelatedLinksMdHeader);
            sb.AppendLine();

            if (_help.RelatedLinks?.Count > 0)
            {
                foreach(var link in _help.RelatedLinks)
                {
                    sb.AppendLine(link.ToRelatedLinksString());
                    sb.AppendLine();
                }
            }
            else
            {
                sb.AppendLine("{{ Fill Related Links Here}}");
                sb.AppendLine();
            }
        }
    }
}
