const puppeteer = require('Puppeteer'),
      cp = require('child_process');


require('dotenv').config();

class PuppeteerDynamics365 {

    constructor(url){
        this._url = url;
    }

    async start(){
        this._browser = await puppeteer.launch({
            headless: false,
            ignoreHTTPSErrors: true
            });
        this._page = await this._browser.newPage();       
        this._page.setViewport({width: 1920, height: 1080});
        this._takeScreenshot = async (fileName, waitForNavigation)=>{
            if(waitForNavigation){
                await this._page.waitForNavigation({waitUntil: 'networkidle', networkIdleTimeout: 2000})
            }
            await this._page.screenshot({path: fileName});
        };
        this._annotateScreenshot = async (label, fileName)=>{
            cp.execSync(`magick -size 600x100 -background blue -font Calibri -pointsize 40 -fill white -gravity center label:"${label}" -bordercolor red -border 8x4 -trim "${fileName}" +swap -gravity south -composite "Annotated-${fileName}"`,
            {cwd: __dirname}); 
        };      
        
        return await this.navigateTo();
    }
    async navigateTo(url){
        await this._page.goto(url || this._url, {waitUntil: 'networkidle'});
        let mainFrame = this._page.mainFrame();
        let childFrames = mainFrame.childFrames();
        //dismiss any popups
        let dialogFrame = childFrames.find(x=>x.name() === 'InlineDialog_Iframe');
        await this._page.evaluate(() => {
          let popup = document.querySelector('#InlineDialog_Iframe');
          if(popup){
            popup.contentDocument.querySelector('#butBegin').click();
          }
        });
        await this._page.click(process.env.TAB_HOME);
        let groups = await this._page.evaluate(() => Array.from(document.querySelectorAll('.nav-group > .navActionButtonContainer')).map(x=>({id: x.id, label: x.querySelector('.navActionButtonLabel').innerText})));
        await this._page.click(process.env.TAB_HOME);
        return groups;
    }
    async navigateToGroup(groups, groupName, screenshotPrefs) {
        let selectedGroup = groups.find(s=>s.label === groupName);
        if(!selectedGroup) {
            Promise.reject(new Error('Not a valid group'));
        }
        let selectedGroupId = `#${selectedGroup.id}`, 
            selectedGroupLabel = selectedGroup.label;
        
        await this._page.click(process.env.TAB_HOME);
        await this._page.click(selectedGroupId, {delay: 500});
        await this._page.click(selectedGroupId, {delay: 500});
        if(screenshotPrefs){
            let fileName = `${screenshotPrefs.fileName || selectedGroup.label}.png`;
            await this._takeScreenshot(fileName);

            if(screenshotPrefs.annotateText){
                await this._annotateScreenshot(screenshotPrefs.annotateText, fileName);
            }
        }
        return await this._page.evaluate(() => Array.from(document.querySelectorAll('.nav-subgroup > .nav-rowBody')).map(x=>({id: x.id, label: x.querySelector('.nav-rowLabel').innerText})));
    }

    async navigateToSubgroup(subgroups, subgroupName, screenshotPrefs){
        let selectedSubgroup = subgroups.find(s=>s.label === subgroupName);
        if(!selectedSubgroup) {
            Promise.reject(new Error('Not a valid subgroup'));
        }
        let selectedSubgroupId = `#${selectedSubgroup.id}`, 
            selectedSubgroupLabel = selectedSubgroup.label;
        await this._page.click(selectedSubgroupId);
        if(screenshotPrefs){
            let fileName = `${screenshotPrefs.fileName || selectedSubgroup.label}.png`;
            await this._takeScreenshot(fileName, true);

            if(screenshotPrefs.annotateText){
                await this._annotateScreenshot(screenshotPrefs.annotateText, fileName);
            }
        }
        await this._page.waitForNavigation({waitUntil: 'networkidle', networkIdleTimeout: 1000});
        return await this._page.evaluate(()=>Array.from(document.querySelectorAll('.ms-crm-CommandBarItem')).map(x=>({id: x.id, label: x.querySelector('a span') ? x.querySelector('a span').innerText : ''})));    
    }

    async clickCommandBarButton(commands, commandName, screenshotPrefs){
        let selectedCommand = commands.find(s=>s.label === commandName);
        if(!selectedCommand) {
            Promise.reject(new Error('Not a valid command bar button'));
        }
        let selectedCommandId = `.ms-crm-CommandBarItem[id="${selectedCommand.id}"] a.ms-crm-Menu-Label`, 
            selectedCommandLabel = selectedCommand.label;
        await this._page.click(selectedCommandId);
        await this._page.waitForNavigation({waitUntil: 'networkidle', networkIdleTimeout: 1000});
        if(screenshotPrefs){
            let fileName = `${screenshotPrefs.fileName || selectedCommand.label}.png`;
            await this._takeScreenshot(fileName, true);

            if(screenshotPrefs.annotateText){
                await this._annotateScreenshot(screenshotPrefs.annotateText, fileName);
            }
        }
    }

    async getAttributes(){
        let attributes = await this._page.evaluate(()=> {
            let xrm = Array.from(Array.from(document.querySelectorAll('iframe')).find(x=>x.style.visibility === 'visible')
            .contentDocument.querySelectorAll('iframe')).find(x=>x.id === 'customScriptsFrame').contentWindow
            .Xrm
            return xrm.Page.getAttribute().map(x=>x.getName());
        });
        return attributes;
    }

    exit(){
        this._browser.close();
    }
}

module.exports = PuppeteerDynamics365;