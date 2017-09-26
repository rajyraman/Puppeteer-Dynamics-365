const fs = require('fs'),
    cp = require('child_process'),
    path = require('path'),
    parse = require('csv-parse');

require('dotenv').config();

if (process.pkg) {
    var puppeteer = require(path.resolve(process.cwd(), 'puppeteer'));
} else {
    var puppeteer = require('puppeteer');
}
let runSheet = fs.readFileSync(process.env.RUN_SHEET_FILE);
let readRunSheetCsv = async () => new Promise(function (resolve, reject) {
    parse(runSheet, { columns: true, trim: true }, (err, rows) => {
        if (err) {
            reject(err);
        }
        else {
            resolve(rows);
        }
    });
});

class PuppeteerDynamics365 {

    constructor(url) {
        this._url = url;
    }

    async start() {
        this._browser = await puppeteer.launch({
            headless: true,
            ignoreHTTPSErrors: true
        });
        this._steps = await readRunSheetCsv();
        this._page = await this._browser.newPage();
        this._page.setViewport({ width: 1920, height: 1080 });

        this._takeScreenshot = async (fileName, waitForNavigation) => {
            if (waitForNavigation) {
                await this._page.waitForNavigation({ waitUntil: 'networkidle', networkIdleTimeout: 2000, networkIdleInflight: 0 });
            }
            await this._page.screenshot({ path: fileName });
        };

        this._annotateScreenshot = async (label, fileName) => {
            cp.execSync(`magick -size 600x100 -background blue -font Calibri -pointsize 40 -fill white -gravity center label:"${label}" -bordercolor red -border 8x4 -trim "${fileName}" +swap -gravity south -composite "${fileName.replace('.', '-Annotated.')}"`);
            await fs.unlink(fileName, c => console.log(`Deleted ${fileName}`));
        };

        this._clickItem = async (rules, screenshotPrefs) => {
            rules = Object.assign({ homeClick: false, waitForNav: true, clickCount: 1 }, rules);
            let selectedItem = rules.items.find(s => s.label === rules.itemName);
            if (!selectedItem) {
                Promise.reject(new Error('Not a valid item'));
            }
            let selectedItemId = rules.selector ? rules.selector.replace('{0}', `${selectedItem.id}`) : `#${selectedItem.id}`,
                selectedItemLabel = selectedItem.label;
            if (rules.homeClick) {
                await this._page.click(process.env.TAB_HOME);
            }
            for (let i = 1; i <= rules.clickCount; i++) {
                try {
                    await this._page.click(selectedItemId, { delay: 500 });
                } catch (e) {
                    console.log(`Error during click of ${selectedItemId}`);
                }
            }
            if (rules.waitForNav) {
                // await this._page.waitFor(4000);
                try {
                    await this._page.waitForNavigation({ timeout: 8000 });
                } catch (e) { }
                //await this._page.waitForNavigation({waitUntil: 'networkidle', networkIdleTimeout: 2000});                                
            }
            if (screenshotPrefs) {
                let fileName = `${screenshotPrefs.fileName || selectedItem.label}.png`;
                await this._takeScreenshot(fileName);

                if (screenshotPrefs.annotateText) {
                    await this._annotateScreenshot(screenshotPrefs.annotateText, fileName);
                }
            }
        };

        return await this.navigateTo();
    }

    async navigateToGroup(groups, groupName, screenshotPrefs) {
        await this._clickItem({
            items: groups,
            itemName: groupName,
            homeClick: true,
            waitForNav: false,
            clickCount: 2
        }, screenshotPrefs);
        return await this._page.evaluate(() => Array.from(document.querySelectorAll('.nav-subgroup > .nav-rowBody')).map(x => ({ id: x.id, label: x.querySelector('.nav-rowLabel').innerText })));
    }

    async navigateToSubgroup(subgroups, subgroupName, screenshotPrefs) {
        let commands = [];
        await this._clickItem({
            items: subgroups,
            itemName: subgroupName,
            clickCount: 2
        }, screenshotPrefs);
        try {
            await this._page.waitForSelector('#crmRibbonManager', { visible: true, timeout: 6000 });
            commands = await this._page.evaluate(() => Array.from(document.querySelectorAll('.ms-crm-CommandBarItem')).map(x => ({ id: x.id, label: x.querySelector('a span') ? x.querySelector('a span').innerText : '' })));
            if (commands.find(c => c.id === 'moreCommands')) {
                await this._clickItem({
                    items: commands,
                    waitForNav: false,
                    itemName: '',
                    selector: '.ms-crm-CommandBarItem[id="moreCommands"] a.ms-crm-Menu-Label'
                });
                commands = await this._page.evaluate(() => Array.from(document.querySelectorAll('.ms-crm-CommandBarItem')).map(x => ({ id: x.id, label: x.querySelector('a span') ? x.querySelector('a span').innerText : '' })));
            }
        } catch (e) {
            console.log(`No visible ribbon for ${subgroupName}`);
        }
        return commands;
    }

    async clickCommandBarButton(commands, commandName, screenshotPrefs) {
        await this._clickItem({
            items: commands,
            itemName: commandName,
            clickCount: 2,
            selector: '.ms-crm-CommandBarItem[id="{0}"] a.ms-crm-Menu-Label'
        }, screenshotPrefs);
    }

    async getAttributes() {
        let attributes = await this._page.evaluate(() => {
            let xrm = Array.from(Array.from(document.querySelectorAll('iframe')).find(x => x.style.visibility === 'visible')
                .contentDocument.querySelectorAll('iframe')).find(x => x.id === 'customScriptsFrame').contentWindow
                .Xrm
            return xrm.Page.getAttribute().map(x => x.getName());
        });
        return attributes;
    }

    async navigateTo(url) {
        await this._page.goto(url || this._url, { waitUntil: 'networkidle' });
        if (process.env.USER_SELECTOR
            && process.env.PASSWORD_SELECTOR
            && process.env.LOGIN_SUBMIT_SELECTOR) {
            await this._page.focus(process.env.USER_SELECTOR);
            await this._page.type(process.env.USER_NAME);
            await this._page.focus(process.env.PASSWORD_SELECTOR);
            await this._page.type(process.env.PASSWORD);
            await this._page.click(process.env.LOGIN_SUBMIT_SELECTOR, { delay: 500 });
            await this._page.waitForNavigation({ waitUntil: 'networkidle', networkIdleTimeout: 2000 });
        }
        let mainFrame = this._page.mainFrame();
        let childFrames = mainFrame.childFrames();
        //dismiss any popups
        let dialogFrame = childFrames.find(x => x.name() === 'InlineDialog_Iframe');
        await this._page.evaluate(() => {
            let popup = document.querySelector('#InlineDialog_Iframe');
            if (popup) {
                popup.contentDocument.querySelector('#butBegin').click();
            }
        });
        await this._page.waitFor(process.env.TAB_HOME);
        await this._page.click(process.env.TAB_HOME);
        let groups = await this._page.evaluate(() => Array.from(document.querySelectorAll('.nav-group > .navActionButtonContainer')).map(x => ({ id: x.id, label: x.querySelector('.navActionButtonLabel').innerText })));
        await this._page.click(process.env.TAB_HOME);
        if (this._steps && this._steps.length > 0) {
            let subgroups = [], commands = [];
            for (const s of this._steps) {
                if (s.group) {
                    console.log(`navigate to group ${s.group}`);
                    subgroups = await this.navigateToGroup(groups, s.group);
                }
                if (s.subgroup) {
                    console.log(`navigate to subgroup ${s.subgroup}`);
                    commands = await this.navigateToSubgroup(subgroups, s.subgroup, { annotateText: s.annotatetext });
                }
                if (s.command) {
                    console.log(`click command ${s.command}`);
                    await this.clickCommandBarButton(commands, s.command, { annotateText: s.annotatetext, fileName: s.filename });
                }
            };
        }
        return groups;
    }

    exit() {
        this._browser.close();
    }
}

module.exports = PuppeteerDynamics365;