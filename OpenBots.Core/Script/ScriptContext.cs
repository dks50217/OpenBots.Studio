﻿using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.CodeAnalysis.Text;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using OBScriptVariable = OpenBots.Core.Script.ScriptVariable;

namespace OpenBots.Core.Script
{
    public class ScriptContext
    {
        public List<OBScriptVariable> Variables { get; set; }
        public List<ScriptArgument> Arguments { get; set; }
        public List<ScriptElement> Elements { get; set; }
        public Dictionary<string, AssemblyReference> ImportedNamespaces { get; set; }
        public List<Assembly> AssembliesList { get; set; }
        public List<string> NamespacesList { get; set; }
        public CSharpCompilationOptions DefaultCompilationOptions { get; set; }
        public List<MetadataReference> DefaultReferences { get; set; }
        public string GuidPlaceholder { get; set; }

        public ScriptContext()
        {
            Variables = new List<OBScriptVariable>();
            Arguments = new List<ScriptArgument>();
            Elements = new List<ScriptElement>();
            ImportedNamespaces = new Dictionary<string, AssemblyReference>(ScriptDefaultNamespaces.DefaultNamespaces);

            AssembliesList = NamespaceMethods.GetAssemblies(ImportedNamespaces);
            NamespacesList = NamespaceMethods.GetNamespacesList(ImportedNamespaces);

            DefaultCompilationOptions = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary).WithOverflowChecks(true)
                                                                                                         .WithOptimizationLevel(OptimizationLevel.Release)
                                                                                                         .WithUsings(NamespacesList);

            DefaultReferences = AssembliesList.Select(x => (MetadataReference)MetadataReference.CreateFromFile(x.Location)).ToList();

            GenerateGuidPlaceHolder();
        }

        public void ReloadCompilerObjects()
        {
            AssembliesList = NamespaceMethods.GetAssemblies(ImportedNamespaces);
            NamespacesList = NamespaceMethods.GetNamespacesList(ImportedNamespaces);

            DefaultCompilationOptions = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary).WithOverflowChecks(true)
                                                                                                         .WithOptimizationLevel(OptimizationLevel.Release)
                                                                                                         .WithUsings(NamespacesList);

            DefaultReferences = AssembliesList.Select(x => (MetadataReference)MetadataReference.CreateFromFile(x.Location)).ToList();
        }

        public void GenerateGuidPlaceHolder()
        {
            GuidPlaceholder = $"v{Guid.NewGuid()}".Replace("-", "");
        }

        public EmitResult EvaluateVariable(string varName, Type varType, string code)
        {
            if (string.IsNullOrEmpty(code))
                code = "null";

            string script = $"{varType.GetRealTypeName()}? {varName} = {code};";

            var parsedSyntaxTree = SyntaxFactory.ParseSyntaxTree(SourceText.From(script, Encoding.UTF8), new CSharpParseOptions(languageVersion: LanguageVersion.CSharp8, kind: SourceCodeKind.Script), "");
            var compilation = CSharpCompilation.Create("CSharp", new SyntaxTree[] { parsedSyntaxTree }, DefaultReferences, DefaultCompilationOptions);
            var result = compilation.Emit("CSharp");

            return result;
        }

        public EmitResult EvaluateInput(Type varType, string code)
        {
            if (string.IsNullOrEmpty(code))
                code = "null";

            var script = "";
            Variables.ForEach(v => script += $"{v.VariableType.GetRealTypeName()}? {v.VariableName} = {(v.VariableValue == null ? "null" : v.VariableValue)};");
            Arguments.ForEach(a => script += $"{a.ArgumentType.GetRealTypeName()}? {a.ArgumentName} = {(a.ArgumentValue == null ? "null" : a.ArgumentValue)};");

            string type;
            var test = varType.GetGenericArguments();
            if (varType.IsGenericType && varType.GetGenericArguments()[0].Name == "T")
                type = "object";
            else
                type = varType.GetRealTypeName();

            GenerateGuidPlaceHolder();
            script += $"{type}? {GuidPlaceholder} = {code};";

            var parsedSyntaxTree = SyntaxFactory.ParseSyntaxTree(SourceText.From(script, Encoding.UTF8), new CSharpParseOptions(languageVersion: LanguageVersion.CSharp8, kind: SourceCodeKind.Script), "");
            var compilation = CSharpCompilation.Create("CSharp", new SyntaxTree[] { parsedSyntaxTree }, DefaultReferences, DefaultCompilationOptions);
            var result = compilation.Emit("CSharp");

            return result;
        }

        public EmitResult EvaluateSnippet(string code)
        {
            if (string.IsNullOrEmpty(code))
                code = "null";

            var script = "";
            Variables.ForEach(v => script += $"{v.VariableType.GetRealTypeName()}? {v.VariableName} = {(v.VariableValue == null ? "null" : v.VariableValue)};");
            Arguments.ForEach(a => script += $"{a.ArgumentType.GetRealTypeName()}? {a.ArgumentName} = {(a.ArgumentValue == null ? "null" : a.ArgumentValue)};");

            script += $"{code};";

            var parsedSyntaxTree = SyntaxFactory.ParseSyntaxTree(SourceText.From(script, Encoding.UTF8), new CSharpParseOptions(languageVersion: LanguageVersion.CSharp8, kind: SourceCodeKind.Script), "");
            var compilation = CSharpCompilation.Create("CSharp", new SyntaxTree[] { parsedSyntaxTree }, DefaultReferences, DefaultCompilationOptions);
            var result = compilation.Emit("CSharp");

            return result;
        }
    }
}