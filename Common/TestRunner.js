var _Om = null;
var _settings = null;
var _item = null;
var _user = null;
var _resultsDictionary = null;
var _executingTests = 0;
var _testLength = 0;

setupGlobals = function () {
    _Om = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    _item = Office.context.mailbox.item;
    _user = Office.context.mailbox.userProfile;
};

// The onclick for the test button.  this will run all tests in the Tests dictionary.
runAllTests = function () {
    _resultsDictionary = {};
    _testLength = Object.keys(sharedCollection).length
    Messages.clear();
    var q = [];
    if (_item.itemType == Office.MailboxEnums.ItemType.Message) {
        _testLength += Object.keys(messageCollection).length;
        Status.reset(_testLength);
        Messages.postMessage("General", "Running Message tests");
        q = createTestQueue(messageCollection);
    }
    else {
        _testLength += Object.keys(meetingCollection).length;
        Status.reset(_testLength);
        Messages.postMessage("General", "Running Meeting tests");
        q = createTestQueue(meetingCollection);
    }

    runTestQueue(q);

    window.setTimeout(checkTests, 30000);
};

createTestQueue = function (specificCollection)
{
    var q = [];

    for (key in specificCollection)
    {
        q.push(specificCollection[key]);
    }

    for (key1 in sharedCollection) {
        q.push(sharedCollection[key1]);
    }

    return q;

}

runTestQueue = function (q)
{
    if (_executingTests == 0)
    {
        var test = q.shift();
        test();
        _executingTests++;
    }
    if (q.length > 0) {
        window.setTimeout(runTestQueue, 1000, q);
    }
}

getAsyncContext = function (name, value, verification) {
    return $_.makeAsyncContext({ testName: name, value: value, callbackVerification: verification });
};

setAsyncContext = function (name) {
    return getAsyncContext(name, null, null);
};

successCallback = function (result) {
    if (Constants.SUCCESS !== result.status) {
        Messages.postMessage(result.asyncContext.testName, "Failed with message: " + result.error.message);
        setResult(result.asyncContext.testName, Constants.FAILURE);
    }
    else {
        try {
            if (result.asyncContext.callbackVerification) {
                result.asyncContext.callbackVerification(result);
            }
            setResult(result.asyncContext.testName, Constants.SUCCESS);
        }
        catch (exception) {
            Messages.postMessage(result.asyncContext.testName, exception);
            setResult(result.asyncContext.testName, Constants.FAILURE);
        }
    }
};

failureCallback = function (result) {
    if (result.status !== Constants.FAILURE) {
        Messages.postMessage(result.asyncContext.testName, "Succeeded, when it was expected to fail.");
        setResult(result.asyncContext.testName, Constants.FAILURE);
    }
    else {
        try {
            if (result.asyncContext.callbackVerification) {
                result.asyncContext.callbackVerification(result);
            }
            setResult(result.asyncContext.testName, Constants.SUCCESS);
        }
        catch (exception) {
            Messages.postMessage(result.asyncContext.testName, exception);
            setResult(result.asyncContext.testName, Constants.FAILURE);
        }
    }
};

skippedCallback = function (result) {
    Messages.postMessage(result.asyncContext.testName, "Skipped");
    setResult(result.asyncContext.testName, Constants.SKIPPED);
};

setResult = function (key, testResult) {
    _executingTests--;
    if (_resultsDictionary[key]) {
        throw "Test " + key + "Tried to set it's result twice, please make sure your test does not call successCallback twice, this will affect the outcome of other tests.";
    }
    else {
        _resultsDictionary[key] = testResult;
    }

    Status.update(Object.keys(_resultsDictionary).length)

    if (Object.keys(_resultsDictionary).length === _testLength) {
        complete();
    }
};
checkTests = function (suite) {
    if (Object.keys(_resultsDictionary).length !== _testLength) {
        for (var key in suite) {
            if (!_resultsDictionary[key]) {
                Messages.postMessage(key, "Timed out");
                setResult(key, Constants.FAILURE);
            }
        }
    }
}
complete = function () {
    Status.complete($_.contains(_resultsDictionary, Constants.FAILURE), $_.contains(_resultsDictionary, Constants.SKIPPED));
};
