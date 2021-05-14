﻿using OpenBots.Commands.Server.HelperMethods;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Engine;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;

namespace OpenBots.Commands.QueueItem.Test
{
    public class SetQueueItemStatusCommandTests
    {
        private AutomationEngineInstance _engine;
        private AddQueueItemCommand _addQueueItem;
        private WorkQueueItemCommand _workQueueItem;
        private SetQueueItemStatusCommand _setQueueItem;

        [Fact]
        public async void SetSuccessfulStatus()
        {
            _engine = new AutomationEngineInstance(null);
            _setQueueItem = new SetQueueItemStatusCommand();

            string name = "SuccessfulQueueItem";
            AddQueueItem(name);
            var queueItemDict = await WorkQueueItem();

            var userInfo = AuthMethods.GetUserInfo();
            var transactionKey = queueItemDict["LockTransactionKey"].ToString();
            var queueItem = QueueItemMethods.GetQueueItemByLockTransactionKey(userInfo.Token, userInfo.ServerUrl, userInfo.OrganizationId, transactionKey);

            VariableMethods.CreateTestVariable(null, _engine, "vQueueItem", typeof(Dictionary<,>));
            _setQueueItem.v_QueueItem = "{vQueueItem}";
            queueItemDict.SetVariableValue(_engine, _setQueueItem.v_QueueItem);
            _setQueueItem.v_QueueItemStatusType = "Successful";

            _setQueueItem.RunCommand(_engine);

            queueItem = QueueItemMethods.GetQueueItemById(userInfo.Token, userInfo.ServerUrl, userInfo.OrganizationId, queueItem.Id);

            Assert.Equal("Success", queueItem.State);
        }

        [Fact]
        public async void SetFailedShouldRetryStatus()
        {
            _engine = new AutomationEngineInstance(null);
            _setQueueItem = new SetQueueItemStatusCommand();

            string name = "NewQueueItem";
            AddQueueItem(name);
            var queueItemDict = await WorkQueueItem();

            var userInfo = AuthMethods.GetUserInfo();
            var transactionKey = queueItemDict["LockTransactionKey"].ToString();
            var queueItem = QueueItemMethods.GetQueueItemByLockTransactionKey(userInfo.Token, userInfo.ServerUrl, userInfo.OrganizationId, transactionKey);

            VariableMethods.CreateTestVariable(null, _engine, "vQueueItem", typeof(Dictionary<,>));
            _setQueueItem.v_QueueItem = "{vQueueItem}";
            queueItemDict.SetVariableValue(_engine, _setQueueItem.v_QueueItem);
            _setQueueItem.v_QueueItemStatusType = "Failed - Should Retry";

            _setQueueItem.RunCommand(_engine);

            queueItem = QueueItemMethods.GetQueueItemById(userInfo.Token, userInfo.ServerUrl, userInfo.OrganizationId, queueItem.Id);

            Assert.Equal("New", queueItem.State);
        }

        [Fact]
        public async void SetFailedFatallyStatus()
        {
            _engine = new AutomationEngineInstance(null);
            _setQueueItem = new SetQueueItemStatusCommand();
            
            string name = "FailedQueueItem";
            AddQueueItem(name);
            var queueItemDict = await WorkQueueItem();

            var userInfo = AuthMethods.GetUserInfo();
            var transactionKey = queueItemDict["LockTransactionKey"].ToString();
            var queueItem = QueueItemMethods.GetQueueItemByLockTransactionKey(userInfo.Token, userInfo.ServerUrl, userInfo.OrganizationId, transactionKey);

            VariableMethods.CreateTestVariable(null, _engine, "vQueueItem", typeof(Dictionary<,>));
            _setQueueItem.v_QueueItem = "{vQueueItem}";
            queueItemDict.SetVariableValue(_engine, _setQueueItem.v_QueueItem);
            _setQueueItem.v_QueueItemStatusType = "Failed - Fatal";

            _setQueueItem.RunCommand(_engine);

            queueItem = QueueItemMethods.GetQueueItemById(userInfo.Token, userInfo.ServerUrl, userInfo.OrganizationId, queueItem.Id);

            Assert.Equal("Failed", queueItem.State);
        }

        [Fact]
        public async Task HandlesNonExistentTransactionKey()
        {
            _engine = new AutomationEngineInstance(null);
            _setQueueItem = new SetQueueItemStatusCommand();

            var queueItemDict = new Dictionary<string, object>()
            {
                { "LockTransactionKey", null },
                { "Name", "ExtendQueueItemTest" },
                { "Source", null },
                { "Event", null },
                { "Type", "Text" },
                { "JsonType", "Test Type" },
                { "DataJson", "Test Text" },
                { "Priority", 10 },
                { "LockedUntilUTC", DateTime.UtcNow.AddHours(1) }
            };

            VariableMethods.CreateTestVariable(null, _engine, "vQueueItem", typeof(Dictionary<,>));
            _setQueueItem.v_QueueItem = "{vQueueItem}";
            _setQueueItem.v_QueueItemStatusType = "Successful";
            queueItemDict.SetVariableValue(_engine, _setQueueItem.v_QueueItem);

            await Assert.ThrowsAsync<NullReferenceException>(() => _setQueueItem.RunCommand(_engine));
        }

        internal void AddQueueItem(string name)
        {
            _addQueueItem = new AddQueueItemCommand();

            _addQueueItem.v_QueueName = "UnitTestQueue";
            _addQueueItem.v_QueueItemName = name;
            _addQueueItem.v_QueueItemType = "Text";
            _addQueueItem.v_JsonType = "Test Type";
            _addQueueItem.v_QueueItemTextValue = "Test Text";
            _addQueueItem.v_Priority = "10";

            _addQueueItem.RunCommand(_engine);
        }

        public async Task<Dictionary<string, object>> WorkQueueItem()
        {
            _workQueueItem = new WorkQueueItemCommand();
            VariableMethods.CreateTestVariable(null, _engine, "output", typeof(Dictionary<,>));

            _workQueueItem.v_QueueName = "UnitTestQueue";
            _workQueueItem.v_OutputUserVariableName = "{output}";
            _workQueueItem.v_SaveAttachments = "No";
            _workQueueItem.v_AttachmentDirectory = "";

            _workQueueItem.RunCommand(_engine);

            var queueItemDict = (Dictionary<string, object>)await "{output}".EvaluateCode(_engine);

            return queueItemDict;
        }
    }
}
