// All tests should follow the signature:
// public bool func();  where the return value is true if the test is synchronous, false if async
// All async objects should pass their test name (key) in userContext.testName

sharedCollection = {

	MakeEwsRequest: function () {
		verify = function (result) {
			Assert.contains("400", result.error.message, "Should be a bad request");
		};
		_Om.makeEwsRequestAsync("JUNK!", failureCallback, { testName: "MakeEwsRequest", callbackVerification: verify });
	},

	// attachment
	addFileAttachment: function () {
		_item.addFileAttachmentAsync(testEmail.fileAttachment.file, testEmail.fileAttachment.name, $_.makeAsyncContext({ testName: "addFileAttachment" }), successCallback)
	},
	addItemAttachment: function () {
		_Om.makeEwsRequestAsync(soapFindItemRequest, addItemAttachmentCallback, { testName: "addItemAttachment" });
	},
	removeFileAttachment: function () {
		_item.removeAttachmentAsync(testEmail.itemAttachmentId, $_.makeAsyncContext({ testName: "removeFileAttachment" }), successCallback);
	},

	// body
	bodySetAsync: function () {
		_item.body.setAsync(testEmail.body, $_.makeAsyncContext({ testName: "bodySetAsync" }, { coercionType: Office.CoercionType.Text }), successCallback);
	},
	bodyGetAsync: function () {
		_item.body.getAsync(Office.CoercionType.Text, $_.makeAsyncContext({ testName: "bodyGetAsync", value: testEmail.body, callbackVerification: verifyObject }), successCallback);
	},
	bodyGetTypeAsync: function () {
		_item.body.getTypeAsync($_.makeAsyncContext({ testName: "bodyGetTypeAsync", value: Office.CoercionType.Html, callbackVerification: verifyObject }), successCallback);
	},
	prependAsync: function () {
		_item.body.prependAsync(testEmail.prependBody, $_.makeAsyncContext({ testName: "prependAsync" }, { coercionType: Office.CoercionType.Text }), successCallback);
	},
	setSelectedDataAsync: function () {
		_item.body.setSelectedDataAsync(testEmail.setSelected, $_.makeAsyncContext({ testName: "setSelectedDataAsync" }, { coercionType: Office.CoercionType.Text }), successCallback);
	},

	// notificationMessages
	notificationMessagesAddAsync: function () {
		_item.notificationMessages.addAsync(testEmail.messageKey, testEmail.firstMessage, setAsyncContext("notificationMessagesAddAsync"), successCallback);
	},
	notificationMessagesReplaceAsync: function () {
		_item.notificationMessages.replaceAsync(testEmail.messageKey, testEmail.secondMessage, setAsyncContext("notificationMessagesReplaceAsync"), successCallback);
	},
	notificationMessagesGetAllAsync: function () {
		_item.notificationMessages.getAllAsync(getAsyncContext("notificationMessagesGetAsync", testEmail.secondMessage, verifyNotification), successCallback);
	},
	notificationMessagesRemoveAsync: function () {
		_item.notificationMessages.removeAsync(testEmail.messageKey, setAsyncContext("notificationMessagesRemoveAsync"), successCallback);
	},

	// subject
	SetSubject: function () {
		_item.subject.setAsync(testEmail.subject, $_.makeAsyncContext({ testName: "SetSubject" }), successCallback)
	},
	GetSubject: function () {
		_item.subject.getAsync($_.makeAsyncContext({ testName: "GetSubject", value: testEmail.subject, callbackVerification: verifyObject }), successCallback)
	},

	saveAsnyc: function () {
		_item.saveAsync($_.makeAsyncContext({ testName: "saveAsnyc" }), successCallback);
	}
};


var messageCollection = {
	// Bcc
	bccAddAsync: function () {
		_item.bcc.addAsync(testEmail.recipients.extra, $_.makeAsyncContext({ testName: "bccAddAsync" }), successCallback);
	},
	bccSetAsync: function () {
		_item.bcc.setAsync(testEmail.recipients.bcc, $_.makeAsyncContext({ testName: "bccSetAsync" }), successCallback);
	},
	bccGetAsync: function () {
		_item.bcc.getAsync($_.makeAsyncContext({ testName: "bccGetAsync", value: testEmail.recipients.bcc, callbackVerification:verifyRecipients }), successCallback);
	},

	// Cc
	ccAddAsync: function () {
		_item.cc.addAsync(testEmail.recipients.extra, $_.makeAsyncContext({ testName: "ccAddAsync" }), successCallback);
	},
	ccSetAsync: function () {
		_item.cc.setAsync(testEmail.recipients.cc, $_.makeAsyncContext({ testName: "ccSetAsync" }), successCallback);
	},
	ccGetAsync: function () {
		_item.cc.getAsync($_.makeAsyncContext({ testName: "ccGetAsync", value: testEmail.recipients.cc, callbackVerification: verifyRecipients }), successCallback);
	},

	// To
	toAddAsync: function () {
		_item.to.addAsync(testEmail.recipients.extra, $_.makeAsyncContext({ testName: "toAddAsync" }), successCallback);
	},
	toSetAsync: function () {
		_item.to.setAsync(testEmail.recipients.to, $_.makeAsyncContext({ testName: "toSetAsync" }), successCallback);
	},
	toGetAsync: function () {
		_item.to.getAsync($_.makeAsyncContext({ testName: "toGetAsync", value: testEmail.recipients.to, callbackVerification: verifyRecipients }), successCallback);
	}
};

var meetingCollection = {


	// Start
	startSetAsync: function () {
		_item.start.setAsync(testEmail.appointment.startTime, $_.makeAsyncContext({ testName: "startSetAsync" }), successCallback);
	},
	startGetAsync: function () {
		_item.start.getAsync($_.makeAsyncContext({ testName: "startGetAsync", value: testEmail.appointment.startTime, callbackVerification: verifyObject }), successCallback);
	},

	// End
	endSetAsync: function () {
		_item.end.setAsync(testEmail.appointment.endTime, $_.makeAsyncContext({ testName: "endSetAsync" }), successCallback);
	},
	endGetAsync: function () {
		_item.end.getAsync($_.makeAsyncContext({ testName: "endGetAsync", value: testEmail.appointment.endTime, callbackVerification: verifyObject }), successCallback);
	},

	// Location
	locationSetAsync: function () {
		_item.location.setAsync(testEmail.appointment.location, $_.makeAsyncContext({ testName: "locationSetAsync" }), successCallback);
	},
	locationGetAsync: function () {
		_item.location.getAsync($_.makeAsyncContext({ testName: "locationGetAsync", value: testEmail.appointment.location, callbackVerification: verifyObject }), successCallback);
	},

	// Optional
	optionalAttendeesAddAsync: function () {
		_item.optionalAttendees.addAsync(testEmail.appointment.extra, $_.makeAsyncContext({ testName: "optionalAttendeesAddAsync" }), successCallback);
	},
	optionalAttendeesSetAsync: function () {
		_item.optionalAttendees.setAsync(testEmail.appointment.optional, $_.makeAsyncContext({ testName: "optionalAttendeesSetAsync" }), successCallback);
	},
	optionalAttendeesGetAsync: function () {
		_item.optionalAttendees.getAsync($_.makeAsyncContext({ testName: "optionalAttendeesGetAsync", value: testEmail.appointment.optional, callbackVerification: verifyRecipients }), successCallback);
	},

	// Required
	requiredAttendeesAddAsync: function () {
		_item.requiredAttendees.addAsync(testEmail.appointment.extra, $_.makeAsyncContext({ testName: "requiredAttendeesAddAsync" }), successCallback);
	},
	requiredAttendeesSetAsync: function () {
		_item.requiredAttendees.setAsync(testEmail.appointment.required, $_.makeAsyncContext({ testName: "requiredAttendeesSetAsync" }), successCallback);
	},
	requiredAttendeesGetAsync: function () {
		_item.requiredAttendees.getAsync($_.makeAsyncContext({ testName: "requiredAttendeesGetAsync", value: testEmail.appointment.required, callbackVerification: verifyRecipients }), successCallback);
	}

};

var messagecount;
jQuery.extend(sharedCollection, commonTestCollection);

addItemAttachmentCallback = function (result) {
	try {
		var parser = new DOMParser();
		var xmlDoc = parser.parseFromString(result.value, "application/xml");
		testEmail.itemAttachmentId = xmlDoc.getElementsByTagName("t:ItemId").length > 0
										? xmlDoc.getElementsByTagName("t:ItemId")[0].attributes.getNamedItem("Id").value
										: xmlDoc.getElementsByTagName("ItemId")[0].attributes.getNamedItem("Id").value;
		Messages.postMessage("General", "Set itemID to " + testEmail.itemAttachmentId);
		_item.addItemAttachmentAsync(testEmail.itemAttachmentId, "Item Attachment", $_.makeAsyncContext({ testName: result.asyncContext.testName }), successCallback);
	}
	catch (e) {
		Messages.postMessage(result.asyncContext.testName, "Failed getItem Request with error: " + e);
		setResult(result.asyncContext.testName, Constants.FAILURE);
	}
}

Office.initialize = function (reason) {
	$(document).ready(function () {
		setupGlobals();
		$("#startButton").click(runAllTests);
	});
}

soapFindItemRequest = "<?xml version='1.0' encoding='utf-8'?>" +
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +
"    xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages' " +
"    xmlns:t='http://schemas.microsoft.com/exchange/services/2006/types' " +
"    xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" +
"  <soap:Header>" +
"    <t:RequestServerVersion Version='Exchange2007_SP1' />" +
"    <t:TimeZoneContext>" +
"      <t:TimeZoneDefinition Id='Eastern Standard Time' />" +
"    </t:TimeZoneContext>" +
"  </soap:Header>" +
"  <soap:Body>" +
"    <m:FindItem Traversal='Shallow'>" +
"      <m:ItemShape>" +
"        <t:BaseShape>IdOnly</t:BaseShape>" +
"      </m:ItemShape>" +
"      <m:IndexedPageItemView MaxEntriesReturned='1' Offset='0' BasePoint='Beginning' />" +
"      <m:ParentFolderIds>" +
"        <t:DistinguishedFolderId Id='inbox' />" +
"      </m:ParentFolderIds>" +
"    </m:FindItem>" +
"  </soap:Body>" +
"</soap:Envelope>"
