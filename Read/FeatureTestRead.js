// A Dictionary that contains all tests
// All tests should follow the signature:
// public bool func();  where the return value is true if the test is synchronous, false if async
// All async objects should pass their test name (key) in userContext.testName

sharedCollection =
{
	// attachments
	attachmentsIsArray: function () {
		var attachments = _item.attachments;

		syncCheck(isArray(attachments), getAsyncContext("attachmentsIsArray", true, verifyObject), successCallback);
	},
	attachmentLength: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments)) {
		    syncCheck(null, setAsyncContext("attachmentLength"), skippedCallback);
		    return;
		}

		syncCheck(attachments.length, getAsyncContext("attachmentLengthIs1", 1, verifyObject), successCallback);
	},
	attachmentType: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments) || attachments.length != 1) {
		    syncCheck(null, setAsyncContext("attachmentIsFile"), skippedCallback);
		    return;
		}

		var attachment = attachments[0];
		syncCheck(attachment.attachmentType, getAsyncContext("attachmentIsFile", AttachmentType.File, verifyObject), successCallback);
	},
	attachmentName: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments) || attachments.length != 1) {
		    syncCheck(null, setAsyncContext("attachmentName"), skippedCallback);
		    return;
		}

		var attachment = attachments[0];
		syncCheck(attachment.name, getAsyncContext("attachmentName", testEmail.attachment.name, verifyObject), successCallback);
	},
	attachmentContentType: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments) || attachments.length != 1) {
		    syncCheck(null, setAsyncContext("attachmentContentType"), skippedCallback);
		    return;
		}

		var attachment = attachments[0];
		syncCheck(!isNullOrEmptyString(attachment.contentType), getAsyncContext("attachmentContentType", true, verifyObject), successCallback);
	},
	attachmentId: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments) || attachments.length != 1) {
		    syncCheck(null, setAsyncContext("attachmentId"), skippedCallback);
		    return;
		}

		var attachment = attachments[0];
		syncCheck(!isNullOrEmptyString(attachment.id), getAsyncContext("attachmentId", true, verifyObject), successCallback);
	},
	attachmentSize: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments) || attachments.length != 1) {
		    syncCheck(null, setAsyncContext("attachmentSize"), skippedCallback);
		    return;
		}

		var attachment = attachments[0];
		syncCheck(!isNullOrEmptyString(attachment.size), getAsyncContext("attachmentSize", true, verifyObject), successCallback);
	},
	attachmentIsInline: function () {
		var attachments = _item.attachments;
		if (!isArray(attachments) || attachments.length != 1) {
		    syncCheck(null, setAsyncContext("attachmentIsInline"), skippedCallback);
		    return;
		}

		var attachment = attachments[0];
		syncCheck(attachment.isInline, getAsyncContext("attachmentIsInline", false, verifyObject), successCallback)
	},

	// Body
	bodyGetAsync: function () {
		var body = testEmail.prependBody + testEmail.body + testEmail.setSelected;
		_item.body.getAsync(Office.CoercionType.Text, $_.makeAsyncContext({ testName: "bodyGetAsync", value: body, callbackVerification: verifyObject }), successCallback);
	},

	// dateTimeCreated
	dateTimeCreated: function () {
		syncCheck(isDate(_item.dateTimeCreated), getAsyncContext("dateTimeCreated", true, verifyObject), successCallback);
	},

	// dateTimeModified
	dateTimeModified: function () {
		syncCheck(isDate(_item.dateTimeModified), getAsyncContext("dateTimeModified", true, verifyObject), successCallback);
	},

	
	// itemId (and also makeEwsRequestAsync - GetItem)
	itemId: function () {
		var itemId = Office.context.mailbox.item.itemId;
		var requestString = getGetItemEwsRequestString(itemId);

		var verify = function (result) {
			var xmlDoc = $.parseXML(result.value);
			$xml = $(xmlDoc);
			$itemId = $xml.find("t\\:ItemId, ItemId");

			if ($itemId.attr("Id") !== result.asyncContext.param) {
				throw "Office.context.mailbox.item.itemId is not equal to the value obtained from EWS.";
			}
		}

		Office.context.mailbox.makeEwsRequestAsync(requestString, successCallback, { testName: "itemId", callbackVerification: verify, param: itemId });

	},

	// normalized subject
	normalizedSubject: function () {
		syncCheck(_item.normalizedSubject, getAsyncContext("normalizedSubject", testEmail.subject, verifyObject), successCallback);
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
	subject: function () {
		syncCheck(_item.subject, getAsyncContext("Subject", testEmail.subject, verifyObject), successCallback);
	},

	GetCallbackTokenAsync: function () {
		var verify = function (result) {
		    if (result.value.length > 200) {
		        throw "Callback token must not be empty";
		    }
		}
		Office.context.mailbox.getCallbackTokenAsync(successCallback, { testName: "GetCallbackTokenAsync", callbackVerification: verify });
	}
};

messageCollection =
{
	// CC
	cc: function () {
		syncCheck(_item.cc, getAsyncContext("cc", testEmail.recipients.cc, verifyRecipients), successCallback);
	},

	// from
	from: function () {
		syncCheck(_item.from.emailAddress, getAsyncContext("from", _user.emailAddress, verifyObject), successCallback);
	},

	// internetMessageId
	internetMessageId: function () {
		syncCheck(!isNullOrEmptyString(_item.internetMessageId), getAsyncContext("internetMessageId", true, verifyObject), successCallback);
	},

	// itemClass
	itemClass: function () {
		syncCheck(isMessageClass(_item.itemClass), getAsyncContext("itemClass", true, verifyObject), successCallback);
	},

	// sender
	sender: function ()
	{
		syncCheck(_item.from.emailAddress, getAsyncContext("sender", _user.emailAddress, verifyObject), successCallback);
	},

	// to
	to: function () {
		syncCheck(_item.to, getAsyncContext("to", testEmail.recipients.to, verifyRecipients), successCallback);
	},

	// displayReplyForm
	displayReplyForm: function () {
		var text = 'hello there from displayReplyForm feature test';
		Office.context.mailbox.item.displayReplyForm(text);
		Messages.postMessage("displayReplyForm", "displayReplyForm has no return. Please confirm reply form shows up with + '" + text + "'");
		syncCheck(null, setAsyncContext("displayReplyForm"), successCallback);
	},
	displayReplyFormWithAttachment: function () {
		var text = 'hello there from displayReplyForm with attachment feature test';
		Office.context.mailbox.item.displayReplyForm(
			{
				'htmlBody': text,
				'attachments':
				[
					{
						'type': Office.MailboxEnums.AttachmentType.File,
						'name': 'Attachment 1',
						'url': 'http://i.imgur.com/ihH1kfg.jpg'
					}
				]
			});

		Messages.postMessage("displayReplyFormWithAttachment", "displayReplyForm has no return. Please manually confirm reply form with attachment shows up with + '" + text + "'");
		syncCheck(null, setAsyncContext("displayReplyFormWithAttachment"), successCallback);
	},
}

meetingCollection =
{
	// end
	end: function () {
		syncCheck(_item.end, getAsyncContext("end", testEmail.appointment.endTime, verifyObject), successCallback);
	},

	// itemClass
	itemClass: function () {
		syncCheck(isAppointmentClass(_item.itemClass), getAsyncContext("itemClass", true, verifyObject), successCallback);
	},

	// location
	location: function () {
		syncCheck(_item.location, getAsyncContext("location", testEmail.appointment.location, verifyObject), successCallback);
	},

	// optionalAttendees
	optionalAttendees: function () {
		syncCheck(_item.optionalAttendees, getAsyncContext("optionalAttendees", testEmail.appointment.optionalAttendees, verifyAttendees), successCallback);
	},

	// organizer
	organizer: function () {
		syncCheck(_item.organizer.emailAddress, getAsyncContext("organizer", _user.emailAddress, verifyObject), successCallback);
	},

	// requiredAttendees
	requiredAttendees: function () {
		syncCheck(_item.requiredAttendees, getAsyncContext("requiredAttendees", testEmail.appointment.requiredAttendees, verifyAttendees), successCallback);
	},

	// start
	start: function () {
		syncCheck(_item.start, getAsyncContext("start", testEmail.appointment.startTime, verifyObject), successCallback);
	},
};

jQuery.extend(sharedCollection, commonTestCollection);

function getGetItemEwsRequestString(itemId) {
	return '<?xml version="1.0"?>' +
	'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
	'<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
	'<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
	'</soap:Header>' +
	'<soap:Body>' +
	'<m:GetItem>' +
	'<m:ItemShape>' +
	'<t:BaseShape>AllProperties</t:BaseShape>' +
	'<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
	'</m:ItemShape>' +
	'<m:ItemIds>' +
	'<t:ItemId Id="' + itemId + '"/>' +
	'</m:ItemIds>' +
	'</m:GetItem>' +
	'</soap:Body>' +
	'</soap:Envelope>';
}

function syncCheck(input, context, callback) {
	context.status = Constants.SUCCESS;
	context.value = input;
	callback(context);
}

Office.initialize = function (reason) {
	$(document).ready(function () {
		setupGlobals();
		$("#startButton").click(runAllTests);
	});
}
