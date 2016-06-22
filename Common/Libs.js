Assert = {
	areEqual: function(expected, actual, text)
	{
		if(expected !== actual)
		{
			throw "{0} Expected: {1}, Actual: {2}".format(text, expected, actual);
		}
	},
	areNotEqual: function(expected, actual, text)
	{
		if(expected === actual)
		{
			throw "{0} Expected: {1}, Actual: {2}".format(text, expected, actual);
		}
	},
	isTrue: function(condition, text)
	{
		if(!condition)
		{
			throw "Expected: False, Actual: True " + text;
		}
	},
	isFalse: function(condition, text)
	{
		if(!!condition)
		{
			throw "Expected: True, Actual: False " + text;
		}
	},
	contains: function(item, collection, text)
	{
		if(!collection.indexOf(item))
		{
			throw "Collection did not contain {1}.  {2}".format(item, text);
		}
	},
	notContains: function(item, collection, text)
	{
		if(collection.indexOf(item))
		{
			throw "Collection did contain {1}.  {2}".format(item, text);
		}
	}
};

verifyRecipients = function (result) {
    var matched = true;
    recipients = [];
    for (i = 0; i < result.value.length; i++) {
        recipients.push(result.value[i].emailAddress);
        if (result.value[i].emailAddress != result.asyncContext.value[i]) {
            matched = false;
        }
    }

    if (!matched) {
        throw "Test " + result.asyncContext.testName + " returned unexpected results, expected '" + result.asyncContext.value + "' got '" + recipients + "'.";
    }
};

verifyObject = function (result) {
    if (result.asyncContext.value.valueOf() !== result.value.valueOf()) {
        throw "Test " + result.asyncContext.testName + " returned unexpected results, expected '" + result.asyncContext.value + "' got '" + result.value + "'.";
    }
};

verifyNotification = function (result) {
    var expected = result.asyncContext.value;
    var got = result.value[0];
    if (expected.type !== got.type || expected.message !== got.message) {
        throw "Test " + result.asyncContext.testName + " returned unexpected results, expected '" + expected.message + "' got '" + got.message + "'.";
    }

};

Constants = {
	SUCCESS: "succeeded",
	FAILURE: "failed",
    SKIPPED: "skipped"
};

Messages = {
	_count: 0,

	postMessage: function(testName, message, color)
	{
		//color = color === undefined ? "rgb(255,0,0)" : color;
		color = this._count % 2 ? "rgb(255,0,255)" : "rgb(100, 0, 100)";
	    var newMessage = "{0} TEST {1}: {2}".format(this._count, testName, message);
		$("#messagesDiv").append("<div style='color:" + color + "'>" + newMessage + "</div><br>\n");
		this._count++;
	},
	clear: function()
	{
		$("#messagesDiv").innerHtml= "";
	}
};

Status = {
	_status: "",
	_totalCount: 0,

	reset: function(totalCount)
	{
		this._totalCount = totalCount;
		this._status = "Ready (0 / {0})".format(this._totalCount);
		this.display("rgb(0,0,0)");
	},
	update: function(count)
	{
		this._status = "Running... ({0} / {1})".format(count, this._totalCount);
		this.display("rgb(255,0,255)");
	},
	complete: function(errorCount, skipped)
	{
		if (!errorCount)
	    {
	    	this._status = "Complete: {0} / {0}, {1} Skipped".format(this._totalCount, skipped)
	        this.display("rgb(0,255,0)")
	    }
	    else
	    {
	        this._status = "Failed: {0} / {1}, {2} Skipped".format(errorCount, this._totalCount, skipped);
	        this.display("rgb(255,0,0)");
	    }

	},
	display: function(color)
	{
		$("#statusDiv").text(this._status);
		$("#statusDiv").css("color", color)
	},
	
};
$_ = {
	// returns number of occurances of value in object
	contains: function(object, value)
	{
		count = 0;
		for( key in object)
			if(object[key] === value)
				count++;
		return count;
	},
	// returns a valid options object with
	makeAsyncContext: function(context, otherObjects)
	{
		if(!otherObjects)
		{
			otherObjects = {};
		}
		otherObjects["asyncContext"] = context;
		return otherObjects;
	}
};

// Check if an item is an appointment.
function itemIsAppointment()
{
	// Source : https://msdn.microsoft.com/en-us/library/5ed24d43-5e45-4c35-8872-c6d3950ad221
	return Office.context.mailbox.item.itemClass === "IPM.Appointment";
}

// Check if an item class is that of an appointment.
function isAppointmentClass(itemClass)
{
	return itemClass === "IPM.Appointment";
}

// Check if an item is a message.
function itemIsMessage()
{
	// Source : https://msdn.microsoft.com/en-us/library/5ed24d43-5e45-4c35-8872-c6d3950ad221
	var itemClass = Office.context.mailbox.item.itemClass;
	
	return itemClass === "IPM.Note"
		|| itemClass === "IPM.Schedule.Meeting.Request"
		|| itemClass === "IPM.Schedule.Meeting.Neg"
		|| itemClass === "IPM.Schedule.Meeting.Pos"
		|| itemClass === "IPM.Schedule.Meeting.Tent"
		|| itemClass === "IPM.Schedule.Meeting.Canceled";
}

// Check if an item class is that of a message.
function isMessageClass(itemClass)
{
	return itemClass === "IPM.Note"
		|| itemClass === "IPM.Schedule.Meeting.Request"
		|| itemClass === "IPM.Schedule.Meeting.Neg"
		|| itemClass === "IPM.Schedule.Meeting.Pos"
		|| itemClass === "IPM.Schedule.Meeting.Tent"
		|| itemClass === "IPM.Schedule.Meeting.Canceled";
}

// Is the object a JavaScript date object or not.
function isDate(obj)
{
	return Object.prototype.toString.call(obj) === '[object Date]';
}

// Is the object a JavaScript array object or not.
function isArray(obj)
{
	return Object.prototype.toString.call(obj) === '[object Array]';
}

function isNullOrEmptyString(str)
{
	if (str && str.toString() !== "") return false;
	return true;
}

// First, checks if it isn't implemented yet.
if (!String.prototype.format) {
  String.prototype.format = function() {
    var args = arguments;
    return this.replace(/{(\d+)}/g, function(match, number) { 
      return typeof args[number] != 'undefined'
        ? args[number]
        : match
      ;
    });
  };
}