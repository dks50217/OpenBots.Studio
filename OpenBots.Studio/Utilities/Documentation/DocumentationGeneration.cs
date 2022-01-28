using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using AContainer = Autofac.IContainer;

namespace OpenBots.Studio.Utilities.Documentation
{
    /// <summary>
    /// 產RPA WEB用的Slides
    /// </summary>
    public class DocumentationGeneration
    {
        /// <summary>
        /// Returns a path that contains the generated markdown files
        /// </summary>
        /// <returns></returns>
        public string GenerateMarkdownFiles(AContainer container, string basePath)
        {
            //create directory if required
            var docsFolderName = "docs2";
            var docsPath = Path.Combine(basePath, docsFolderName);
            if (!Directory.Exists(docsPath))
            {
                Directory.CreateDirectory(docsPath);
            }

            var commandClasses = TypeMethods.GenerateCommandTypes(container);

            var highLevelCommandInfo = new List<CommandMetaData>();
            StringBuilder stringBuilder = null;
            string fullFileName = String.Empty;

            var groupList = commandClasses.GroupBy(c => c.Namespace).Select(c => c.Key);

            foreach (var markdownName in groupList)
            {
                // Create Folder
                string kebobDestination = markdownName.Replace(".", "_").Replace(" ", "-").Replace("/", "-").ToLower();
                var destinationdirectory = docsPath;//Path.Combine(docsPath, kebobDestination);
                var kebobFileName = $"{kebobDestination.Replace("openbots_commands_", string.Empty)}-commands.md";

                stringBuilder = new StringBuilder();

                if (!Directory.Exists(destinationdirectory))
                {
                    Directory.CreateDirectory(destinationdirectory);
                }

                // Header Style
                var styleHeader = "<style> .small td{font-size:40%;}th{font-size:70%} .label foreignObject {overflow: visible;} svg[id^='mermaid - '] { width:400% } </style>";
                stringBuilder.AppendLine(styleHeader);
                stringBuilder.AppendLine(Environment.NewLine);

                // Command Name
                stringBuilder.AppendLine("# " + kebobDestination);

                // Header Menu
                stringBuilder.AppendLine();
                stringBuilder.AppendLine("---");

                // Empty Menu Title
                stringBuilder.AppendLine();
                stringBuilder.AppendLine("#");
                stringBuilder.AppendLine();

                stringBuilder.AppendLine("<div class='table-wrapper small' markdown='block'>");
                stringBuilder.AppendLine();
                stringBuilder.AppendLine(@"| Commands                  | Sub\_Commands                    |");
                stringBuilder.AppendLine("| ------------------------  |:--------------------------------:|");

                var filterClass = commandClasses.Where(c=> c.FullName.Contains(markdownName)).ToList();

                foreach (var commandClass in filterClass)
                {
                    ScriptCommand instantiatedCommand = (ScriptCommand)Activator.CreateInstance(commandClass);
                    var commandName = instantiatedCommand.SelectionName;
                    stringBuilder.AppendLine($"| {markdownName}               | {commandName}                             |");
                }

                stringBuilder.AppendLine();
                stringBuilder.AppendLine("</div>");

                stringBuilder.AppendLine();
                stringBuilder.AppendLine("---");

                //loop each command
                foreach (var commandClass in filterClass)
                {
                    //instantiate and pull properties from command class
                    ScriptCommand instantiatedCommand = (ScriptCommand)Activator.CreateInstance(commandClass);
                    var groupName = GetClassValue(commandClass, typeof(CategoryAttribute));
                    var classDescription = GetClassValue(commandClass, typeof(DescriptionAttribute));
                    var commandName = instantiatedCommand.SelectionName;

                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("#### " + commandName);
                    stringBuilder.AppendLine();

                    // slide table
                    stringBuilder.AppendLine("<div class='table-wrapper small' markdown='block'>");
                    stringBuilder.AppendLine();

                    stringBuilder.AppendLine("| Parameter Question   	| What to input  	|  Sample Data 	| Remarks  	|");
                    stringBuilder.AppendLine("| ---                    | ---               | ---           | ---       |");

                    //loop each property
                    foreach (var prop in commandClass.GetProperties().Where(f => f.Name.StartsWith("v_")).ToList())
                    {
                        //pull attributes from property
                        var commandLabel = CleanMarkdownValue(GetPropertyValue(prop, typeof(DisplayNameAttribute)));
                        var helpfulExplanation = CleanMarkdownValue(GetPropertyValue(prop, typeof(DescriptionAttribute)));
                        var sampleUsage = CleanMarkdownValue(GetPropertyValue(prop, typeof(SampleUsage)));
                        var remarks = CleanMarkdownValue(GetPropertyValue(prop, typeof(Remarks)));

                        //append to parameter table
                        stringBuilder.AppendLine("|" + commandLabel + "|" + helpfulExplanation + "|" + sampleUsage + "|" + remarks + "|");
                    }

                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("</div>");
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("---");

                    //write file
                    fullFileName = Path.Combine(destinationdirectory, kebobFileName);
                }

                File.WriteAllText(fullFileName, stringBuilder.ToString());
                stringBuilder.Clear();
            }

            return docsPath;
        }

        /// <summary>
        /// 產RPA WEB用的JSON FILE
        /// </summary>
        /// <param name="container"></param>
        /// <param name="basePath"></param>
        /// <returns></returns>
        public string GenerateJsonFiles(AContainer container, string basePath)
        {
            //create directory if required
            //var docsFolderName = "json";
            //var docsPath = Path.Combine(basePath, docsFolderName);
            //if (!Directory.Exists(docsPath))
            //{
            //    Directory.CreateDirectory(docsPath);
            //}

            //var commandClasses = TypeMethods.GenerateCommandTypes(container);
            return String.Empty;
        }

        private string GetPropertyValue(PropertyInfo prop, Type attributeType)
        {
            var attribute = prop.GetCustomAttributes(attributeType, true);

            if (attribute.Length == 0)
            {
                return "Data not specified";
            }
            else
            {
                var attributeFound = attribute[0];

                if (attributeFound is DisplayNameAttribute)
                {
                    var processedAttribute = (DisplayNameAttribute)attributeFound;
                    return processedAttribute.DisplayName;
                }
                else if (attributeFound is DescriptionAttribute)
                {
                    var processedAttribute = (DescriptionAttribute)attributeFound;
                    return processedAttribute.Description;
                }
                else if (attributeFound is SampleUsage)
                {
                    var processedAttribute = (SampleUsage)attributeFound;
                    return processedAttribute.Usage;
                }
                else if (attributeFound is Remarks)
                {
                    var processedAttribute = (Remarks)attributeFound;
                    return processedAttribute.Remark;
                }
                else
                {
                    return "Attribute not supported";
                }
            }
        }

        private string CleanMarkdownValue(string value)
        {
            Dictionary<string, string> replacementDict = new Dictionary<string, string>
            {
                {"|", "\\|"},
                {"\n\t", "<br>"},
                {"\r\n", "<br>"}
            };
            foreach (var replacementTuple in replacementDict)
            {
                value = value.Replace(replacementTuple.Key, replacementTuple.Value);
            }

            return value;
        }

        private string GetClassValue(Type commandClass, Type attributeType)
        {
            var attribute = commandClass.GetCustomAttributes(attributeType, true);

            if (attribute.Length == 0)
            {
                return "Data not specified";
            }
            else
            {
                var attributeFound = attribute[0];

                if (attributeFound is CategoryAttribute)
                {
                    var processedAttribute = (CategoryAttribute)attributeFound;
                    return processedAttribute.Category;
                }
                else if (attributeFound is DescriptionAttribute)
                {
                    var processedAttribute = (DescriptionAttribute)attributeFound;
                    return processedAttribute.Description;
                }
            }

            //string groupAttribute = "";
            //if (attribute.Length > 0)
            //{
            //    var attributeFound = (attributeType)attribute[0];
            //    groupAttribute = attributeFound.groupName;
            //}

            return "OK";
        }
    }
}
