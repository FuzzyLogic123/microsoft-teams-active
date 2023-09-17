const requestURL = 'https://presence.teams.microsoft.com/v1/me/reportmyactivity/'
let activityRequest;

const getActivityRequest = async () => {
    const puppeteer = require('puppeteer');

    await (async () => {
        const browser = await puppeteer.launch({ headless: false, userDataDir: './user_data' });
        const page = await browser.newPage();

        await page.setRequestInterception(true);

        page.on('request', (request) => {
            if (request.url() === requestURL && request.method() === 'POST') {
                activityRequest = {
                    url: request.url(),
                    method: request.method(),
                    headers: request.headers(),
                    body: request.postData(),
                };
            }
            request.continue();
        });

        await page.goto('https://teams.microsoft.com/');
        await page.waitForRequest(request => {
            return request.url() === requestURL && request.method() === 'POST';
        }, {
            visible: true,
            waitUntil: "load",
            waitUntil: "networkidle0",
            waitUntil: "domcontentloaded",
            waitUntil: "networkidle2",
            timeout: 600 * 1000,
        });
        await browser.close();
    })();
}

const makeRequests = async (activityRequest) => {
    activityRequest.body = JSON.stringify({ ...JSON.parse(activityRequest.body), isActive: true });
    const interval = setInterval(() => {
        if (new Date().getHours() <= 16) {
            fetch(activityRequest.url, activityRequest);
        } else {
            clearInterval(interval);
        }
    }, 1000 * 60 * 5)
}

async function main() {
    await getActivityRequest();
    makeRequests(activityRequest);
}

main();
