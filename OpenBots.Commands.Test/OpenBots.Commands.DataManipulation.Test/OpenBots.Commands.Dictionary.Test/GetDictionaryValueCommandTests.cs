﻿using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Engine;
using System.Collections.Generic;
using Xunit;

namespace OpenBots.Commands.Dictionary.Test
{
    public class GetDictionaryValueCommandTests
    {
        private GetDictionaryValueCommand _getDictionaryValue;
        private AutomationEngineInstance _engine;

        [Fact]
        public async void GetsDictionaryValue()
        {
            _getDictionaryValue = new GetDictionaryValueCommand();
            _engine = new AutomationEngineInstance(null);

            Dictionary<string, string> inputDictionary = new Dictionary<string, string>();
            inputDictionary.Add("key1", "val1");
            VariableMethods.CreateTestVariable(inputDictionary, _engine, "inputDictionary", typeof(Dictionary<,>));
            VariableMethods.CreateTestVariable(null, _engine, "outputValue", typeof(string));

            _getDictionaryValue.v_InputDictionary = "{inputDictionary}";
            _getDictionaryValue.v_Key = "key1";
            _getDictionaryValue.v_OutputUserVariableName = "{outputValue}";

            _getDictionaryValue.RunCommand(_engine);

            Assert.Equal("val1", (string)await "{outputValue}".EvaluateCode(_engine));
        }
    }
}
