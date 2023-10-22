const requestURL = 'https://presence.teams.microsoft.com/v1/me/reportmyactivity/'
let activityRequest;

const getActivityRequest = async () => {
    const puppeteer = require('puppeteer');

    await (async () => {
        const browser = await puppeteer.launch({ headless: true, userDataDir: './user_data' });
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

const makeRequests = async () => {
    activityRequest.body = JSON.stringify({ ...JSON.parse(activityRequest.body), isActive: true });
    const interval = setInterval(async () => {
        if (new Date().getHours() <= 16) {
            const result = await fetch(activityRequest.url, activityRequest);
            const contentType = result.headers.get('Content-Type');
            if (contentType && contentType.includes('text')) {
                console.log(result.text());
            }

            if (result.status !== 200) {
                console.log("Re-authorizing")
                try {
                    await getActivityRequest();
                } catch (e) {
                    console.log(e)
                }
            }
        } else {
            console.log("After 5pm, stopping requests")
            clearInterval(interval);
        }
    }, 1000 * 60 * 1)
}

async function main() {
    await getActivityRequest();
    makeRequests();
}

main();