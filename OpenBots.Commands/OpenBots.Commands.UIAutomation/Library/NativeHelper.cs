﻿using OpenBots.Core.ChromeNativeClient;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace OpenBots.Commands.UIAutomation.Library
{
    public class NativeHelper
    {
        public async static Task<WebElement> DataTableToWebElement(DataTable SearchParametersDT, IAutomationEngineInstance engine) 
        {
			WebElement webElement = new WebElement();
			webElement.XPath = (SearchParametersDT.Rows[0].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[0].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.RelXPath = (SearchParametersDT.Rows[1].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[1].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.ID = (SearchParametersDT.Rows[2].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[2].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.Name = (SearchParametersDT.Rows[3].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[3].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.TagName = (SearchParametersDT.Rows[4].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[4].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.ClassName = (SearchParametersDT.Rows[5].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[5].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.LinkText = (SearchParametersDT.Rows[6].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[6].ItemArray[2].ToString().EvaluateCode(engine) : "";
			webElement.CssSelector = (SearchParametersDT.Rows[7].ItemArray[0].ToString().ToLower() == "true") ?
				(string)await SearchParametersDT.Rows[7].ItemArray[2].ToString().EvaluateCode(engine) : "";
			return webElement;
		}

		public static DataTable WebElementToDataTable(WebElement webElement)
		{
			DataTable SearchParameters = NewSearchParameterDataTable();
			SearchParameters.Rows.Add(true, "\"XPath\"", $"\"{webElement.XPath}\"");
			SearchParameters.Rows.Add(true, "\"Relative XPath\"", $"\"{webElement.RelXPath}\"");
			SearchParameters.Rows.Add(false, "\"ID\"", $"\"{webElement.ID}\"");
			SearchParameters.Rows.Add(false, "\"Name\"", $"\"{webElement.Name}\"");
			SearchParameters.Rows.Add(false, "\"Tag Name\"", $"\"{webElement.TagName}\"");
			SearchParameters.Rows.Add(false, "\"Class Name\"", $"\"{webElement.ClassName}\"");
			SearchParameters.Rows.Add(false, "\"Link Text\"", $"\"{webElement.LinkText}\"");
			SearchParameters.Rows.Add(false, "\"CSS Selector\"", $"\"{webElement.CssSelector}\"");
			return SearchParameters;
		}

		public static DataTable WebElementToSeleniumDataTable(WebElement webElement)
		{
			DataTable SearchParameters = NewSearchParameterDataTable();
			SearchParameters.Rows.Add(true, "\"XPath\"", $"\"{webElement.XPath}\"");
			SearchParameters.Rows.Add(true, "\"Relative XPath\"", $"\"{webElement.RelXPath}\"");
			SearchParameters.Rows.Add(false, "\"ID\"", $"\"{webElement.ID}\"");
			SearchParameters.Rows.Add(false, "\"Name\"", $"\"{webElement.Name}\"");
			SearchParameters.Rows.Add(false, "\"Tag Name\"", $"\"{webElement.TagName}\"");
			SearchParameters.Rows.Add(false, "\"Class Name\"", $"\"{webElement.ClassName}\"");
			SearchParameters.Rows.Add(false, "\"Link Text\"", $"\"{webElement.LinkText}\"");
			var _cssSelectors = webElement.CssSelectors.Split(',').ToList();
			for (int i = 0; i < _cssSelectors.Count; i++)
				SearchParameters.Rows.Add(false, $"\"CSS Selector {i + 1}\"", $"\"{_cssSelectors[i]}\"");
			return SearchParameters;
		}

		public static DataTable NewSearchParameterDataTable()
		{
			DataTable searchParameters = new DataTable();
			searchParameters.Columns.Add("Enabled");
			searchParameters.Columns.Add("Parameter Name");
			searchParameters.Columns.Add("Parameter Value");
			searchParameters.TableName = DateTime.Now.ToString("UIASearchParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"));
			return searchParameters;
		}
    }
}
