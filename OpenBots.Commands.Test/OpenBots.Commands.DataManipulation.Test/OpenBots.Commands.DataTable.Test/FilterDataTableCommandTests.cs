﻿using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Engine;
using System.Data;
using Xunit;
using Xunit.Abstractions;
using OBData = System.Data;

namespace OpenBots.Commands.DataTable.Test
{
    public class FilterDataTableCommandTests
    {

        private FilterDataTableCommand _filterDataTable;
        private AutomationEngineInstance _engine;
        private readonly ITestOutputHelper output;

        public FilterDataTableCommandTests(ITestOutputHelper output)
        {
            this.output = output;
        }
        [Fact]
        public async void filtersDataTable()
        {
            _filterDataTable = new FilterDataTableCommand();
            _engine = new AutomationEngineInstance(null);

            OBData.DataTable tableToFilter = new OBData.DataTable();
            tableToFilter.Columns.Add("col1");
            tableToFilter.Columns.Add("col2");
            DataRow row1 = tableToFilter.NewRow();
            row1["col1"] = "id1";
            row1["col2"] = "data1";
            tableToFilter.Rows.Add(row1);
            DataRow row2 = tableToFilter.NewRow();
            row2["col1"] = "id2";
            row2["col2"] = "data2";
            tableToFilter.Rows.Add(row2);

            VariableMethods.CreateTestVariable("col1", _engine, "col1", typeof(string));
            VariableMethods.CreateTestVariable("id1", _engine, "id1", typeof(string));
            VariableMethods.CreateTestVariable(tableToFilter, _engine, "tableToFilter", typeof(OBData.DataTable));
            VariableMethods.CreateTestVariable(null, _engine, "outputTable", typeof(OBData.DataTable));

            _filterDataTable.v_DataTable = "{tableToFilter}";
            //_filterDataTable.v_SearchItem = "({col1},{id1})";
            _filterDataTable.v_OutputUserVariableName = "{outputTable}";

            _filterDataTable.RunCommand(_engine);

            OBData.DataTable expectedDT = new OBData.DataTable();
            expectedDT.Columns.Add("col1");
            expectedDT.Columns.Add("col2");
            DataRow row1copy = expectedDT.NewRow();
            row1copy["col1"] = "id1";
            row1copy["col2"] = "data1";
            expectedDT.Rows.Add(row1copy);
            
            OBData.DataTable resultDataTable = (OBData.DataTable)await _filterDataTable.v_OutputUserVariableName.EvaluateCode(_engine);
            // Check each row / column pair and assert equivalence
            output.WriteLine(expectedDT.Rows[0].ToString());
            output.WriteLine(resultDataTable.Rows[0].ToString());
            for (int row = 0; row < expectedDT.Rows.Count; row++)
            {
                for (int col = 0; col < expectedDT.Columns.Count; col++)
                {
                    Assert.Equal(expectedDT.Rows[row][expectedDT.Columns[col]], resultDataTable.Rows[row][resultDataTable.Columns[col]]);
                }
            }


        }
    }
}
