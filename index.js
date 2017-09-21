const PuppeteerDynamics365 = require('./puppeteerdynamics365');

require('dotenv').config();

(async () => {
  const pup = new PuppeteerDynamics365(process.env.CRM_URL);
  try{
    let groups = await pup.start();
    let subgroups = await pup.navigateToGroup(groups, 'Workspace', {annotateText:'Workspace'});
    let commands = await pup.navigateToSubgroup(subgroups, 'Clients', {annotateText:'Clients view'});
    console.log(commands);
    await pup.clickCommandBarButton(commands, 'NEW', {annotateText:'Create new client', fileName: 'Create new client'});
    let attributes = await pup.getAttributes();
    console.log(attributes);
    subgroups = await pup.navigateToGroup(groups, 'Settings');
    commands = await pup.navigateToSubgroup(subgroups, 'Administration', {annotateText:'Administration area'}); 
    
    subgroups = await pup.navigateToGroup(groups, 'Marketing');
    await pup.navigateToSubgroup(subgroups, 'Leads', {annotateText:'View all existing active leads'});
      
  }catch(e){
    console.log(e);
  }  
  pup.exit();
})();