let gleanGlobals = {};

self.addEventListener('message',
    function (event) {
        let job = event.data;
        if (typeof (job) == "string") {
            job = JSON.parse(job);
        }

        const jobId = job.jobId;
        try {
            const result = invokeFunction(job.name, job.parameters);
            // check whether the result is a promise.
            if (typeof (result) == "function" || typeof (result) == "object" && typeof (result.then) == "function") {
                result.then(function (realResult) {
                    postMessage(
                        {
                            jobId: jobId,
                            result: realResult
                        }
                    );
                })
                    .catch(function (ex) {
                        postMessage(
                            {
                                jobId: jobId,
                                error: true
                            }
                        )
                    });
            }
            else {
                postMessage({
                    jobId: jobId,
                    result: result
                });
            }
        }
        catch (ex) {
            postMessage({
                jobId: jobId,
                error: true
            });
        }
        if (event.data && event.data.type === "init") {
            gleanGlobals = event.data.gleanGlobals;
            return;
        }
    }
);

function invokeFunction(name, parameters) {
    if (name == "SEARCH") {
        return webRequest.apply(null, parameters);
    }
    else {
        throw new Error("not supported");
    }
}

function webRequest(input) {
    let url = "https://" + gleanGlobals.instance + "-be.glean.com/rest/api/v1/chat";
    
    console.log("Glean Instance: ", gleanGlobals.instance);

    let headers = {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + gleanGlobals.token,
    };

    let data = {
        'stream': false,
        'messages': [{
            'author': 'USER',
            'fragments': [{ 'text': `${input}` }]
        }]
    };

    return new Promise(function (resolve, reject) {
        fetch(url, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(data),
            muteHttpExceptions: true
        })
            .then(function (response) {
                if (response.status === 200) {
                    return response.json();
                } else {
                    throw new Error(`Error: ${response.status}, ${response.statusText}`);
                }
            })
            .then(function (jsonResponse) {
                // Process the response using the same logic as processFinalContent
                let messages = jsonResponse.messages || [];
                let finalContent = '';

                messages.forEach(function (message) {
                    let messageType = message.messageType;
                    let fragments = message.fragments || [];

                    if (messageType === 'CONTENT') {
                        // Concatenate fragments
                        fragments.forEach(function (fragment) {
                            finalContent += fragment.text || '';
                        });
                    }
                });

                resolve(finalContent);
            })
            .catch(function (error) {
                console.log('Request Exception: ' + error.message);
                reject('Error: Request failed - ' + error.message);
            });
    });
}