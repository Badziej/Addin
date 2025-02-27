// Copyright © 2021 Forcepoint LLC. All rights reserved.

let logEnable = false;

function sleep(delay) {
    var start = new Date().getTime();
    while (new Date().getTime() < start + delay);
}

Office.initialize = function () {
    Office.onReady().then(function() {
        if (operatingSytem() !== "MacOS") {
            console.error("This add-in is restricted to macOS. Exiting...");
            return; // Stop further execution
        }
        printLog("Add-in initialized on MacOS");
    });
};

function printLog(text) {
    console.log(text);
    if(logEnable && (typeof text === 'string' || text instanceof String)) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("succeeded", {
            type: "progressIndicator",
            message: text.substring(0, Math.min(text.length, 250)),
        });
        sleep(1500);
    }
}

async function httpServerCheck(resolve, reject){
    printLog("Checking the server");
    const controller = new AbortController();
    const timeout = setTimeout(() => {
        controller.abort();
    }, 30000);

    fetch('https://localhost:55296/FirefoxExt/_1', {
        method: 'GET',
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'same-origin',
        redirect: 'follow',
        referrerPolicy: 'no-referrer',
    }).then(response => {
        clearTimeout(timeout);
        if (!response.ok) {
            printLog("Server is down");
            reject(false);
        }
        printLog("Server is UP");
        resolve(true);
    }).catch(e => {
        printLog(e);
        printLog("Request crashed");
        reject(false);
    });
}

async function sendToClasifier(url = '', data = {}, event) {
    printLog("Sending event to classifier");
    const controller = new AbortController();
    const timeout = setTimeout(() => {
        controller.abort();
    }, 30000);

    fetch(url, {
        signal: controller.signal,
        method: 'POST',
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'same-origin',
        headers: {
            'Content-Type': 'application/json'
        },
        redirect: 'follow',
        referrerPolicy: 'no-referrer',
        body: JSON.stringify(data)
    }).then(response => {
        if (!response.ok) {
            printLog("Engine returned error: "+response.json());
            handleError(response, event);
        }
        return response.json();
    }).then(response => {
        clearTimeout(timeout);
        handleResponse(response, event);
    }).catch(e => {
        printLog(e);
        printLog("Request crashed");
        handleError(e, event);
    });
}

function handleResponse(data, event) {
    printLog("Handling response from engine");
    let message = Office.context.mailbox.item;
    if (data["action"] === 1) {
        message.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked by DLP engine' });
        printLog("DLP block");
        event.completed({ allowEvent: false });
    } else {
        printLog("DLP allow");
        event.completed({ allowEvent: true });
    }
}

function operatingSytem() { 
    var contextInfo = Office.context.diagnostics;
    printLog('Office application: ' + contextInfo.host);
    printLog('Platform: ' + contextInfo.platform);
    printLog('Office version: ' + contextInfo.version);

    if (contextInfo.platform === 'Mac') {
        return 'MacOS';
    }
    return 'Other';
}

function validateBody(event) {
    Office.onReady().then(function() {
        printLog("FP email validation started - [v1.2]");

        // Execute the add-in logic only if it is Outlook running on MacOS
        if (operatingSytem() !== "MacOS") {
            printLog("OS is not MacOS. Blocking execution.");
            handleError("Not MacOS", event);
            return;
        }

        printLog("MacOS detected");
        validate(event).catch(data => { handleError(data, event); });
    });
}

function handleError(data, event) {
    printLog(data);
    printLog(event);
    printLog("Completing event ");
    event.completed({ allowEvent: true });
    printLog("Event Completed");
}

// Prevents UI elements from displaying on non-MacOS devices
document.addEventListener("DOMContentLoaded", function() {
    if (operatingSytem() !== "MacOS") {
        document.body.innerHTML = "<h2>This add-in is only available on macOS.</h2>";
    }
});

if (typeof exports !== 'undefined') {
    exports.handleResponse = handleResponse;
    exports.handleError = handleError;
    exports.validateBody = validateBody;
}
