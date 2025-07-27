const GleanNamespace = {};

(function(GleanNamespace) {
    // The max number of web workers to be created
    const g_maxWebWorkers = 4;

    // The array of web workers
    const g_webworkers = [];
    
    // Next job id
    let g_nextJobId = 0;

    // The promise info for the job. It stores the {resolve: resolve, reject: reject} information for the job.
    const g_jobIdToPromiseInfoMap = {};

    function getOrCreateWebWorker(jobId) {
        const index = jobId % g_maxWebWorkers;
        if (g_webworkers[index]) {
            return g_webworkers[index];
        }

        // create a new web worker
        const webWorker = new Worker("functions-worker.js");
        webWorker.addEventListener('message', function(event) {
            let jobResult = event.data;
            if (typeof(jobResult) == "string") {
                jobResult = JSON.parse(jobResult);
            }

            if (typeof(jobResult.jobId) == "number") {
                const jobId = jobResult.jobId;
                // get the promise info associated with the job id
                const promiseInfo = g_jobIdToPromiseInfoMap[jobId];
                if (promiseInfo) {
                    if (jobResult.error) {
                        // The web worker returned an error
                        promiseInfo.reject(new Error());
                    }
                    else {
                        // The web worker returned a result
                        promiseInfo.resolve(jobResult.result);
                    }
                    delete g_jobIdToPromiseInfoMap[jobId];
                }
            }
        });

        webWorker.postMessage({
            type: "init",
            gleanGlobals: window.gleanGlobals
        });

        g_webworkers[index] = webWorker;
        return webWorker;
    }

    function dispatchSearchJob(functionName, parameters) {
        const jobId = g_nextJobId++;
        return new Promise(function(resolve, reject) {
            // store the promise information.
            g_jobIdToPromiseInfoMap[jobId] = {resolve: resolve, reject: reject};
            const worker = getOrCreateWebWorker(jobId);
            worker.postMessage({
                jobId: jobId,
                name: functionName,
                parameters: parameters
            });
        });
    }

    GleanNamespace.dispatchSearchJob = dispatchSearchJob;
})(GleanNamespace);

CustomFunctions.associate("SEARCH", function(n) {
    return GleanNamespace.dispatchSearchJob("SEARCH", [n]);
});
