"use strict";

const base64 = require('base-64');
const fetch = require('node-fetch');
const path = require('path');
const fs = require('fs');

const consolePrefix = 'devops-testrun-connect: ';
let runs = new Map();
let basicauth;
let devopsRootURL;
let isReady = false;
let ownerName = null;


// Wrap fetch for devops API calls
async function callAPI(endpoint, method, body = null) {
    let b = body === null? null : JSON.stringify(body);
    return fetch(endpoint, {
        method: method,
        body: b,
        headers: {
            'content-type': 'application/json',
            authorization : basicauth,
        }
    });
}

exports.init = function (subscription, project, pat, owner) {
    if (subscription && project && pat) {
        basicauth = `Basic ${base64.encode(`:${pat}`)}`;
        devopsRootURL = `https://dev.azure.com/${subscription}/${project}/_apis/test/`;
        isReady = true;
        ownerName = owner;
    } else {
        console.warn(`${consolePrefix}cannot initialise the module. <subscription:${subscription}>, <project:${project}>, <pat:${pat}>`);
    }
}

exports.startRun = async function (planID, suiteID, name) {
    if (!isReady) {
        console.warn(`${consolePrefix}cannot start run before module is initialised. Use 'init(subscription, project, pat)' first.`);
        return undefined;
    } 

    // Retrieve Test Suite name
    let suiteName = '';
    await callAPI(  `${devopsRootURL}Plans/${planID}/suites/${suiteID}?api-version=5.0`, 
                    'GET'
    )
    .then(response => response.json())
    .then(data => { 
        suiteName = data.name;
     })
    .catch((error) => {
        console.error(error);
    });

    // Retrieve Test Cases in Test Suite
    let testcases = [];
    await callAPI(  `${devopsRootURL}Plans/${planID}/suites/${suiteID}/testcases?api-version=5.0`, 
                    'GET'
    )
    .then(response => response.json())
    .then(data => { 
        for (let elem of data.value) {
            testcases.push(elem.testCase.id);
        }
     })
    .catch((error) => {
        console.error(error);
    });
    
    console.log(`${consolePrefix}found test cases: ${testcases}`);

    // Retrieve Test Points
    let runpoints = [];
    let testpoints = [];
    await callAPI(  `${devopsRootURL}points?api-version=5.0-preview.2`, 
                    'POST', 
                    {
                        "PointsFilter": {
                            "TestcaseIds": testcases
                        }
                    }
    )
    .then(response => response.json())
    .then(data => { 
        for (let elem of data.points) {
            if (elem.testPlan.id === planID && elem.suite.id === suiteID) {
                runpoints.push({tc:elem.testCase.id, tp:elem.id, url:elem.url});
            }
        }
     })
    .catch((error) => {
        console.error(error);
    });

    testpoints = runpoints.map((x) => x.tp);
    console.log(`${consolePrefix}found test points:${testpoints}`);
    
    // Create Test Run containing Test Points
    let runID = undefined;
    name = (name !== undefined) ? name : `Automated <${suiteName}> (ID:${suiteID})`;
    await callAPI(  `${devopsRootURL}runs?api-version=5.0`,
                    'POST',
                    {
                        "name" : name,
                        "automated" : true,
                        "plan" : {
                            "id" : planID
                        },
                        "pointIds" : testpoints,
                        "owner" : {
                            "displayName" : ownerName
                        }
                    }
    )
    .then(response => response.json())
    .then(data => {        
        runID = data.id;
        console.log(`${consolePrefix}test run created (${runID})`);
    })
    .catch((error) => {
        console.error(error);
    });

    if (runID === undefined) {
        console.log(`${consolePrefix}failed creating new test run.. progress won't be visible in devops`);
        return;
    }

    // Retrieve Test Results and all other IDs
    let tests = [];
    await callAPI(  `${devopsRootURL}Runs/${runID}/results?api-version=5.0`,
                    'GET'
    )
    .then(response => response.json())
    .then(data => {      
        for (let value of data.value) {
            let test = {
                "id": value.testCase.id,
                "point": value.testPoint.id,
                "result": value.id,
                "startTime": 0
            }
            tests.push(test);
        }
    })
    .catch((error) => {
        console.error(error);
    });
    
    // store for later
    runs.set(runID, tests);

    return runID;
}

exports.startTest = async function (runID, testID) {
    if (!isReady) {
        console.warn(`${consolePrefix}cannot start before module is initialised. Use 'init(subscription, project, pat)' first.`);
    } 

    let tests = runs.get(runID);
    let test = tests.find(x => x.id === testID); 

    if (test === undefined) {
        console.log(`${consolePrefix}not test point found for test case ${testID}`);
        return;
    }

    await callAPI(  `${devopsRootURL}runs/${runID}/results?api-version=5.0`,
                    'PATCH',
                    [{
                        "id" : test.result,
                        "testPoint" : { "id" : test.point },
                        "state" : "InProgress",
                        "owner" : {
                            "displayName" : ownerName
                        }
                    }]
    )
    .then(response => response.json())
    .then(data => { 
        test.startTime = Date.now();
        //console.log(`${consolePrefix}updated outcome for ${test.id}`);
    })
    .catch((error) => {
        console.error(error);
    });
}

exports.endTest = async function (runID, testID, outcome) {
    if (!isReady) {
        console.warn(`${consolePrefix}cannot end test before module is initialised. Use 'init(subscription, project, pat)' first.`);
    } 

    let tests = runs.get(runID);
    let test = tests.find(x => x.id === testID); 

    if (test === undefined) {
        console.log(`${consolePrefix}not test point found for test case ${testID}`);
        return;
    }

    let elapsed = Date.now() - test.startTime;

    console.log(`ELAPSED ${elapsed}`);

    await callAPI(  `${devopsRootURL}runs/${runID}/results?api-version=6.0`,
                    'PATCH',
                    [{
                        "id" : test.result,
                        "testPoint" : { "id" : test.point },
                        "outcome" : outcome,
                        "state" : "Completed",
                        "durationInMs" : elapsed
                    }]
    )
    .then(response => response.json())
    .then(data => { 
        //console.log(`${consolePrefix}updated outcome for ${test.id}`);
    })
    .catch((error) => {
        console.error(error);
    });
}

exports.endRun = async function (runID, reportPath = undefined) {
    if (!isReady) {
        console.warn(`${consolePrefix}cannot end run before module is initialised. Use 'init(subscription, project, pat)' first.`);
    } 

    // Update run state to completed
    await callAPI(  `${devopsRootURL}runs/${runID}?api-version=5.0`,
                    'PATCH',
                    {
                        "state":"Completed"
                    }
    )
    .then(response => response.json())
    .then(data => {         
        console.log(`${consolePrefix}test run ended (${runID})`);
    })
    .catch((error) => {
        console.error(error);
    });

    // Upload report if provided
    if (reportPath !== undefined) {
        const content = fs.readFileSync(reportPath, {encoding: 'base64'});
        await callAPI(  `${devopsRootURL}runs/${runID}/attachments?api-version=5.0-preview.1`,
                        'POST',
                        {
                            "stream": content,
                            "fileName": path.parse(reportPath).base,
                            "comment": "Test Report",
                            "attachmentType": "GeneralAttachment"
                        }
        )
        .then(response => response.json())
        .then(data => {         
            console.log(`${consolePrefix}report attached`);
        })
        .catch((error) => {
            console.error(error);
        });
    }
}

