const puppeteer = require('Puppeteer'),
cp = require('child_process');


require('dotenv').config();

class PuppeteerDynamics365 {

constructor(url) {
    this._url = url;
}

async start() {
    this._browser = await puppeteer.launch({
        headless: false,
        ignoreHTTPSErrors: true
    });
    this._page = await this._browser.newPage();
    this._page.setViewport({ width: 1920, height: 1080 });

    this._takeScreenshot = async (fileName, waitForNavigation) => {
        if (waitForNavigation) {
            await this._page.waitForNavigation({ waitUntil: 'networkidle', networkIdleTimeout: 2000 });
        }
        await this._page.screenshot({ path: fileName });
    };

    this._annotateScreenshot = async (label, fileName) => {
        cp.execSync(`magick -size 600x100 -background blue -font Calibri -pointsize 40 -fill white -gravity center label:"${label}" -bordercolor red -border 8x4 -trim "${fileName}" +swap -gravity south -composite "Annotated-${fileName}"`,
            { cwd: __dirname });
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
            await this._page.click(selectedItemId, { delay: 500 });
        }
        if (rules.waitForNav) {
            // await this._page.waitFor(4000);
            try{
                await this._page.waitForNavigation({timeout: 6000});
            }catch(e){}
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
async navigateTo(url) {
    await this._page.goto(url || this._url, { waitUntil: 'networkidle' });
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
    await this._page.click(process.env.TAB_HOME);
    let groups = await this._page.evaluate(() => Array.from(document.querySelectorAll('.nav-group > .navActionButtonContainer')).map(x => ({ id: x.id, label: x.querySelector('.navActionButtonLabel').innerText })));
    await this._page.click(process.env.TAB_HOME);
    return groups;
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
    try{
        await this._page.waitForSelector('#crmRibbonManager',{visible: true, timeout: 6000});
        commands = await this._page.evaluate(() => Array.from(document.querySelectorAll('.ms-crm-CommandBarItem')).map(x => ({ id: x.id, label: x.querySelector('a span') ? x.querySelector('a span').innerText : '' })));
        if(commands.find(c => c.id === 'moreCommands')){
            await this._clickItem({
                items: commands,
                waitForNav: false,
                itemName: '',
                selector: '.ms-crm-CommandBarItem[id="moreCommands"] a.ms-crm-Menu-Label'
            });
            commands = await this._page.evaluate(() => Array.from(document.querySelectorAll('.ms-crm-CommandBarItem')).map(x => ({ id: x.id, label: x.querySelector('a span') ? x.querySelector('a span').innerText : '' })));
        }            
    }catch(e){
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

exit() {
    this._browser.close();
}
}

module.exports = PuppeteerDynamics365;