const PuppeteerDynamics365 = require('./puppeteerdynamics365');

require('dotenv').config();

(async () => {
  let missingSettings = ['CRM_URL','TAB_HOME','USER_SELECTOR','USER_NAME','PASSWORD_SELECTOR','PASSWORD','LOGIN_SUBMIT_SELECTOR','RUN_SHEET_FILE']
  .filter(s=>!process.env[s]);
  if(missingSettings.length > 0){
    console.log('Check if the .env file exists and the following settings exist\r\n---------------------------------');
    console.log(missingSettings);
    return;
  }   
  const pup = new PuppeteerDynamics365(process.env.CRM_URL);
  try{
    let groups = await pup.start();
    // let subgroups = await pup.navigateToGroup(groups, 'Settings');
    // let commands = await pup.navigateToSubgroup(subgroups, 'Administration', {annotateText:'Administration area'}); 
    
    // subgroups = await pup.navigateToGroup(groups, 'Marketing');
    // commands = await pup.navigateToSubgroup(subgroups, 'Leads', {annotateText:'View all existing active leads'});

    // await pup.clickCommandBarButton(commands, 'NEW', {annotateText:'Create new lead', fileName: 'Create new lead'});
    // let attributes = await pup.getAttributes();
    // console.log(attributes);      
  }catch(e){
    console.log(e);
  }  
  pup.exit();
})();