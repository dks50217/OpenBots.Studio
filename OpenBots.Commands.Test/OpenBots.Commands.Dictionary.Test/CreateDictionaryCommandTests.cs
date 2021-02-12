﻿using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Engine;
using System.Collections.Generic;
using System.Data;
using Xunit;
using OBData = System.Data;

namespace OpenBots.Commands.Dictionary.Test
{
    public class CreateDictionaryCommandTests
    {
        private CreateDictionaryCommand _createDictionary;
        private AutomationEngineInstance _engine;

        [Fact]
        public void CreatesDictionary()
        {
            _createDictionary = new CreateDictionaryCommand();
            _engine = new AutomationEngineInstance(null);
            OBData.DataTable inputDt = new OBData.DataTable();
            inputDt.Columns.Add("Keys");
            inputDt.Columns.Add("Values");
            DataRow row1 = inputDt.NewRow();
            row1["Keys"] = "key1";
            row1["Values"] = "val1";
            inputDt.Rows.Add(row1);
            inputDt.CreateTestVariable(_engine, "inputDt");
            "unassigned".CreateTestVariable(_engine, "output");

            _createDictionary.v_ColumnNameDataTable = (OBData.DataTable)"{inputDt}".ConvertUserVariableToObject(_engine);
            _createDictionary.v_OutputUserVariableName = "{output}";

            _createDictionary.RunCommand(_engine);

            Dictionary<string, string> outDict = (Dictionary<string, string>)"{output}".ConvertUserVariableToObject(_engine);

            Assert.True(outDict.ContainsKey("key1"));
            Assert.Equal("val1", outDict["key1"]);
        }
    }
}
