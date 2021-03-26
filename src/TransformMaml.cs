using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.PowerShell.PlatyPS.Model;

namespace Microsoft.PowerShell.PlatyPS
{
    internal class TransformMaml : TransformBase
    {
        public TransformMaml(PSSession session) : base(session)
        {
        }

        internal override Collection<CommandHelp> Transform(string[] mamlFileNames)
        {
            Collection<CommandHelp> cmdHelp = new();

            foreach (var file in mamlFileNames)
            {
                if (!File.Exists(file))
                {
                    throw new ArgumentNullException($"File {file} does not exist");
                }

                foreach (var command in ReadMaml(file))
                {
                    cmdHelp.Add(command);
                }
            }

            return cmdHelp;
        }

        private Collection<CommandHelp> ReadMaml(string mamlFile)
        {
            Collection<CommandHelp> commandHelps = new();

            using StreamReader stream = new(mamlFile);

            XmlReaderSettings settings = new();

            using XmlReader reader = XmlReader.Create(stream, settings);

            if (reader.MoveToContent() != XmlNodeType.Element)
            {
                throw new InvalidOperationException(reader.NodeType.ToString() + "is invalid XmlNode");
            }

            while (reader.ReadToFollowing(Constants.MamlCommandCommandTag))
            {
                commandHelps.Add(ReadCommand(reader.ReadSubtree()));
            }

            return commandHelps;
        }

        private CommandHelp ReadCommand(XmlReader reader)
        {
            CommandHelp cmdHelp = new();

            if (reader.ReadToFollowing(Constants.MamlCommandNameTag))
            {
                cmdHelp.Title = reader.ReadElementContentAsString();
                cmdHelp.Synopsis = ReadSynopsis(reader);
                cmdHelp.Description = ReadDescription(reader);
                cmdHelp.AddSyntaxItemRange(ReadSyntaxItems(reader));
            }
            else
            {
                throw new InvalidOperationException("maml file is invalid");
            }

            return cmdHelp;
        }

        private Collection<SyntaxItem> ReadSyntaxItems(XmlReader reader)
        {
            Collection<SyntaxItem> items = new();

            if (reader.ReadToFollowing(Constants.MamlSyntaxTag))
            {
                if(reader.ReadToDescendant(Constants.MamlSyntaxItemTag))
                {
                    do
                    {
                        items.Add(ReadSyntaxItem(reader.ReadSubtree()));

                        // needed to go to next command:syntaxitem
                        reader.MoveToElement();

                    } while (reader.ReadToNextSibling(Constants.MamlSyntaxItemTag));
                }
            }

            return items;
        }

        private SyntaxItem ReadSyntaxItem(XmlReader reader)
        {
            if (reader.ReadToDescendant(Constants.MamlNameTag))
            {
                string commandName = reader.ReadElementContentAsString();

                //Collection<CommandInfo> cmdInfo = PowerShellAPI.GetCommandInfo(commandName);

                // @TODO Get parameter set name info
                SyntaxItem syntaxItem = new SyntaxItem(commandName, parameterSetName: null, isDefaultParameterSet: false);

                while (reader.ReadToNextSibling(Constants.MamlCommandParameterTag))
                {
                    Parameter parameter = new Parameter();

                    if (reader.HasAttributes)
                    {
                        if (reader.MoveToAttribute("required"))
                        {
                            bool required;
                            if (bool.TryParse(reader.Value, out required))
                            {
                                parameter.Required = required;
                            }
                        }

                        if (reader.MoveToAttribute("variableLength"))
                        {
                            bool variableLength;
                            if (bool.TryParse(reader.Value, out variableLength))
                            {
                                parameter.VariableLength = variableLength;
                            }
                        }

                        if (reader.MoveToAttribute("globbing"))
                        {
                            bool globbing;
                            if (bool.TryParse(reader.Value, out globbing))
                            {
                                parameter.Globbing = globbing;
                            }
                        }

                        if (reader.MoveToAttribute("pipelineInput"))
                        {
                            // Value is like 'True (ByPropertyName, ByValue)' or 'False'
                            parameter.PipelineInput = reader.Value.StartsWith(Constants.TrueString, StringComparison.OrdinalIgnoreCase) ? true : false;
                        }

                        if (reader.MoveToAttribute("position"))
                        {
                            // Value is like '0' or 'named'
                            parameter.Position = reader.Value;
                        }

                        if (reader.MoveToAttribute("aliases"))
                        {
                            parameter.Aliases = reader.Value;
                        }

                        reader.MoveToElement();
                    }

                    if (reader.ReadToDescendant(Constants.MamlNameTag))
                    {
                        parameter.Name = reader.ReadElementContentAsString();
                    }

                    parameter.Description = ReadDescription(reader);

                    if (reader.Read())
                    {
                        if (string.Equals(reader.Name, Constants.MamlCommandParameterValueGroupTag, StringComparison.OrdinalIgnoreCase))
                        {
                            if (reader.ReadToDescendant(Constants.MamlCommandParameterValueTag))
                            {
                                do
                                {
                                    parameter.AddAcceptedValue(reader.ReadElementContentAsString());
                                } while (reader.ReadToNextSibling(Constants.MamlCommandParameterValueTag));
                            }
                        }
                        /*else if (string.Equals(reader.Name, Constants.MamlCommandParameterValueTag, StringComparison.OrdinalIgnoreCase))
                        {
                            //parameter.Type = Type.GetType(reader.ReadElementContentAsString());
                            reader.MoveToElement();
                            reader.ReadEndElement();
                        }*/
                    }

                    if (reader.ReadToNextSibling(Constants.MamlDevTypeTag))
                    {
                        if (reader.ReadToDescendant(Constants.MamlNameTag))
                        {
                            parameter.Type = Type.GetType(reader.ReadElementContentAsString());
                        }
                    }

                    if (reader.ReadToFollowing(Constants.MamlDevDefaultValueTag))
                    {
                        parameter.DefaultValue = reader.ReadElementContentAsString();
                    }

                    syntaxItem.AddParameter(parameter);

                    // need to go the end of command:parameter
                    if (reader.ReadState != ReadState.EndOfFile)
                    {
                        reader.ReadEndElement();
                    }
                }

                return syntaxItem;
            }

            return null;
        }

        private string ReadSynopsis(XmlReader reader)
        {
            string synopsis = null;

            if (reader.ReadToNextSibling(Constants.MamlDescriptionTag))
            {
                if (reader.ReadToDescendant(Constants.MamlParaTag))
                {
                    synopsis = reader.ReadElementContentAsString();
                }
            }

            return synopsis;
        }

        private string ReadDescription(XmlReader reader)
        {
            StringBuilder description = new();

            if (reader.ReadToFollowing(Constants.MamlDescriptionTag))
            {
                if (reader.ReadToDescendant(Constants.MamlParaTag))
                {
                    do
                    {
                        description.AppendLine(reader.ReadElementContentAsString());
                        description.AppendLine();
                    } while (reader.ReadToNextSibling(Constants.MamlParaTag));
                }
            }

            reader.ReadEndElement();

            return description.ToString().TrimEnd(Environment.NewLine.ToCharArray());
        }
    }
}
