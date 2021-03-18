using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using Microsoft.PowerShell.PlatyPS.Model;

namespace Microsoft.PowerShell.PlatyPS
{
    internal class TransformCommand : TransformBase
    {
        public TransformCommand(PSSession session) : base(session)
        {
        }

        internal override Collection<CommandHelp> Transform(string[] commandNames)
        {
            Collection<CommandHelp> cmdHelp = new();

            foreach (var command in commandNames)
            {
                Collection<CommandInfo> cmdletInfos = PowerShellAPI.GetCommandInfo(command);

                foreach (var cmdletInfo in cmdletInfos)
                {
                    cmdHelp.Add(ConvertCmdletInfo(cmdletInfo));
                }
            }

            return cmdHelp;
        }

        private CommandHelp ConvertCmdletInfo(CommandInfo cmdletInfo)
        {
            Collection<PSObject> help = PowerShellAPI.GetHelpForCmdlet(cmdletInfo.Name);

            bool addDefaultStrings = false;
            dynamic helpItem = null;

            if (help?.Count == 1)
            {
                helpItem = help[0];

                // If the description and examples are empty the help is auto-generated.
                // So assume that no existing help content is available.
                if (string.IsNullOrEmpty(helpItem.description) &&
                    string.IsNullOrEmpty(helpItem.examples))
                {
                    addDefaultStrings = true;
                }
            }
            else
            {
                addDefaultStrings = true;
            }

            CommandHelp cmdHelp = new();

            cmdHelp.Title = cmdletInfo.Name;
            cmdHelp.ModuleName = cmdletInfo.ModuleName;
            cmdHelp.Synopsis = GetSynopsis(helpItem, addDefaultStrings);
            cmdHelp.AddSyntaxItemRange(GetSyntaxItem(cmdletInfo, helpItem));
            cmdHelp.Description = GetDescription(helpItem, addDefaultStrings);
            cmdHelp.AddExampleItemRange(GetExamples(helpItem, addDefaultStrings));
            cmdHelp.AddParameterRange(GetParameters(cmdletInfo, helpItem, addDefaultStrings));

            cmdHelp.AddInputItem(
                GetInputOutputItem(
                    helpItem.inputTypes.inputType,
                    addDefaultStrings ? Constants.NoneString : null,
                    addDefaultStrings ? string.Empty : null)); ;

            cmdHelp.AddOutputItem(
                GetInputOutputItem(
                    helpItem.returnValues.returnValue,
                    addDefaultStrings ? Constants.SystemObjectTypename : null,
                    addDefaultStrings ? string.Empty : null));

            cmdHelp.Notes = GetNotes(helpItem, addDefaultStrings);
            cmdHelp.AddReleatedLinksRange(GetRelatedLinks(helpItem));

            return cmdHelp;
        }

        private static IEnumerable<Parameter> GetParameters(CommandInfo cmdletInfo, dynamic helpItem, bool addDefaultString)
        {
            List<Parameter> parameters = new();

            foreach (KeyValuePair<string, ParameterMetadata> parameterMetadata in cmdletInfo.Parameters)
            {
                Parameter param = new();
                param.Name = parameterMetadata.Value.Name;

                foreach (KeyValuePair<string, ParameterSetMetadata> paramSet in parameterMetadata.Value.ParameterSets)
                {
                    string paramSetName = paramSet.Key;
                    param.AddParameterSet(paramSetName);
                    param.AddRequiredParameterSets(paramSet.Value.IsMandatory, paramSetName);
                }

                param.DefaultValue = GetParameterDefaultValueFromHelp(helpItem, param.Name);
                param.Aliases = string.Join(",", parameterMetadata.Value.Aliases);

                string descriptionFromHelp = GetParameterDescriptionFromHelp(helpItem, param.Name);
                param.Description = string.IsNullOrEmpty(descriptionFromHelp) ?
                    string.Format(Constants.FillInParameterDescriptionTemplate, param.Name) :
                    descriptionFromHelp;

                param.Type = parameterMetadata.Value.ParameterType;

                GetParameterAtributeInfo(parameterMetadata.Value.Attributes, ref param);

                parameters.Add(param);
            }

            return parameters;
        }

        private static IEnumerable<Example> GetExamples(dynamic helpItem, bool addDefaultString)
        {
            List<Example> examples = new();

            if (addDefaultString)
            {
                Example exp = new();
                exp.Title = "Example 1";
                exp.Code = Constants.FillInExampleCode;
                exp.Remarks = Constants.FillInExampleDescription;
                examples.Add(exp);
            }
            else
            {
                int exampleCounter = 1;

                var examplesArray = helpItem?.examples?.example;

                if (examplesArray != null)
                {
                    foreach (var item in examplesArray)
                    {
                        Example exp = new();
                        exp.Code = item.code.ToString();
                        exp.Remarks = GetStringFromDescriptionArray(item.remarks);
                        string title = item.title.ToString().Trim(' ', '-').Replace($"Example {exampleCounter}: ", "");
                        exp.Title = title;

                        examples.Add(exp);
                    }
                }
            }

            return examples;
        }

        private static List<Links> GetRelatedLinks(dynamic helpItem)
        {
            List<Links> links = new();

            if (helpItem?.relatedLinks?.navigationLink != null)
            {
                foreach (dynamic navlink in helpItem.relatedLinks.navigationLink)
                {
                    links.Add(new Links(navlink.uri.ToString(), navlink.linkText.ToString()));

                }
            }

            return links;
        }

        private static IEnumerable<SyntaxItem> GetSyntaxItem(CommandInfo cmdletInfo, dynamic helpItem)
        {
            List<SyntaxItem> syntaxItems = new();

            foreach (CommandParameterSetInfo parameterSetInfo in cmdletInfo.ParameterSets)
            {
                SyntaxItem syn = new(cmdletInfo.Name, parameterSetInfo.Name, parameterSetInfo.IsDefault);

                foreach (CommandParameterInfo paramInfo in parameterSetInfo.Parameters)
                {
                    Parameter param = GetParameterInfo(cmdletInfo, helpItem, paramInfo);
                    syn.AddParameter(param);
                }

                syntaxItems.Add(syn);
            }

            return syntaxItems;
        }

        private static Parameter GetParameterInfo(CommandInfo cmdletInfo, dynamic helpItem, CommandParameterInfo paramInfo)
        {
            Parameter param = new Parameter();

            string paramName = paramInfo.Name;

            param.Name = paramName;

            string descriptionFromHelp = GetParameterDescriptionFromHelp(helpItem, paramName);
            param.Description = string.IsNullOrEmpty(descriptionFromHelp) ?
                string.Format(Constants.FillInParameterDescriptionTemplate, paramName) :
                descriptionFromHelp;

            param.Type = paramInfo.ParameterType;

            param.AddParameterSetsRange(GetParameterSetsOfParameter(paramName, cmdletInfo));

            param.Aliases = string.Join("-", paramInfo.Aliases);
            param.Required = paramInfo.IsMandatory;

            param.Position = paramInfo.Position == int.MinValue ? Constants.NamedString : paramInfo.Position.ToString();

            string defaultValueFromHelp = GetParameterDefaultValueFromHelp(helpItem, paramName);
            param.DefaultValue = string.IsNullOrEmpty(defaultValueFromHelp) ?
                Constants.NoneString :
                defaultValueFromHelp;

            param.PipelineInput = paramInfo.ValueFromPipeline | paramInfo.ValueFromPipelineByPropertyName;

            GetParameterAtributeInfo(paramInfo.Attributes, ref param);

            return param;
        }

        private static void GetParameterAtributeInfo(IEnumerable<Attribute> attributes, ref Parameter param)
        {
            foreach (var attrib in attributes)
            {
                switch (attrib)
                {
                    case ParameterAttribute parameterAttribute:
                        param.DontShow = parameterAttribute.DontShow;
                        param.PipelineInput = parameterAttribute.ValueFromPipeline | parameterAttribute.ValueFromPipelineByPropertyName;
                        param.Position = parameterAttribute.Position == int.MinValue ? Constants.NamedString : parameterAttribute.Position.ToString();
                        param.Required = parameterAttribute.Mandatory;
                        break;

                    case SupportsWildcardsAttribute wildcardsAttribute:
                        param.Globbing = true;
                        break;

                    case ValidateSetAttribute validateSetAttribute:
                        param.AddAcceptedValues(validateSetAttribute.ValidValues);
                        break;
                }
            }
        }

        private static IEnumerable<string> GetParameterSetsOfParameter(string parameterName, CommandInfo cmdletInfo)
        {
            if (cmdletInfo.Parameters.TryGetValue(parameterName, out ParameterMetadata paramMetadata))
            {
                return paramMetadata.ParameterSets.Keys;
            }

            return null;
        }

        private static string GetParameterDescriptionFromHelp(dynamic helpItem, string parameterName)
        {
            foreach (dynamic parameter in helpItem.parameters.parameter)
            {
                if (string.Equals(parameter.name.ToString(), parameterName, StringComparison.OrdinalIgnoreCase))
                {
                    return GetStringFromDescriptionArray(parameter.description);
                }
            }

            return null;
        }

        private static string GetParameterDefaultValueFromHelp(dynamic helpItem, string parameterName)
        {
            foreach (dynamic parameter in helpItem.parameters.parameter)
            {
                if (string.Equals(parameter.name.ToString(), parameterName, StringComparison.OrdinalIgnoreCase))
                {
                    return parameter?.defaultValue?.ToString();
                }
            }

            return null;
        }

        private static string GetNotes(dynamic helpItem, bool addDefaultString)
        {
            if (addDefaultString)
            {
                return Constants.FillInNotes;
            }
            else
            {
                return helpItem?.alertSet?.alert != null ?
                    GetStringFromDescriptionArray(helpItem.alertSet.alert) :
                    null;
            }
        }

        private static string GetDescription(dynamic helpItem, bool addDefaultStrings)
        {
            if (addDefaultStrings)
            {
                return Constants.FillInDescription;
            }
            else
            {
                if (helpItem == null)
                {
                    throw new ArgumentNullException(nameof(helpItem));
                }

                return GetStringFromDescriptionArray(helpItem.description);
            }
        }

        private static string GetSynopsis(dynamic helpItem, bool addDefaultStrings)
        {
            if (addDefaultStrings)
            {
                return Constants.FillInSynopsis;
            }
            else
            {
                return helpItem != null ? helpItem.Synopsis.ToString() : throw new ArgumentNullException(nameof(helpItem));
            }
        }

        private static InputOutput GetInputOutputItem(dynamic typesInfo, string defaultTypeName, string defaultDescription)
        {
            InputOutput inputOutputTypeItem = new();

            if (string.IsNullOrEmpty(defaultTypeName) && string.IsNullOrEmpty(defaultDescription))
            {
                dynamic ioTypes = typesInfo;

                if (ioTypes is IEnumerable<PSObject>)
                {
                    foreach (dynamic ioType in typesInfo)
                    {
                        inputOutputTypeItem.AddInputOutputItem(ioType.type.name.ToString(), GetStringFromDescriptionArray(ioType.description));
                    }
                }
                else if (ioTypes is PSObject)
                {
                    inputOutputTypeItem.AddInputOutputItem(ioTypes.type.name.ToString(), GetStringFromDescriptionArray(ioTypes.description));
                }
            }
            else
            {
                inputOutputTypeItem.AddInputOutputItem(defaultTypeName, defaultDescription);
            }

            return inputOutputTypeItem;
        }

        private static string GetStringFromDescriptionArray(dynamic description)
        {
            if (description == null)
            {
                return null;
            }

            StringBuilder sb = new();

            foreach (dynamic line in description)
            {
                string text = line.text.ToString();

                // Add semantic line break.
                sb.AppendLine(text.Replace(". ", $".{Environment.NewLine}"));
            }

            return sb.ToString();
        }
    }
}