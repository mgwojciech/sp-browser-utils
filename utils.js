function debugTab() {
    let tabFrame = document.querySelector("embedded-page-container iframe[name='embedded-page-container']")
    tabFrame.src += encodeURIComponent("&debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js");
}
function debugPage() {
    let param = "";
    if (window.location.href.indexOf("?") >= 0) {
        param = "&"
    } else {
        param = "?"
    }
    window.location.href += (param + "debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js");
}


function openConfig() {
    let tabFrame = document.querySelector("embedded-page-container iframe[name='embedded-page-container']")
    tabFrame.src += encodeURIComponent("&openPropertyPane=true");
}

function checkIfApp() {
    let queryParams = new URLSearchParams(window.location.search);
    let listId = queryParams.get("list");
    return get(`/sites/tea-point/_api/web/lists(guid'${listId}')?$select=EntityTypeName`).then(data => {
        return data.json().then(result => result.EntityTypeName === "HostedAppConfigsList");
    })
}
function debugWebPart(webPartId) {
    let tabFrame = document.querySelector("embedded-page-container iframe[name='embedded-page-container']")
    let tabSrc = new URL(tabFrame.src);
    tabFrame.src = `${tabSrc.origin}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=${webPartId}%26forceLocale=en-us${encodeURIComponent("&debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js")}`
}
function get(api){
    return fetch(api,{
        method:"GET",
        headers:{
            accept: "application/json",
            'content-type': "application/json;odata=verbose"
        }
    })
}

function getVerbose(api){
    return fetch(api,{
        method:"GET",
        headers:{
            accept: "application/json;odata=verbose",
            'content-type': "application/json;odata=verbose"
        }
    })
}
function getRequestDigest(url){
    return fetch(url + "/_api/contextinfo",{method: "POST", headers:{accept:"application/json"}}).then(result=>{ return result.json().then(json=>{console.log(json); return json;})});
}

function post(siteUrl, api, updateBody){
    getRequestDigest(siteUrl).then(digestResponse=>{
       let digest = digestResponse.FormDigestValue;
       fetch(siteUrl + api,{
           method:"POST",
           headers:{
               accept: "application/json;odata=nometadata",
               "X-RequestDigest": digest,
               'content-type': "application/json"
           },
           body: JSON.stringify(updateBody)
       })
    });
}
const permissionKind = {
    "emptyMask": 0,
    "viewListItems": 1,
    "addListItems": 2,
    "editListItems": 3,
    "deleteListItems": 4,
    "approveItems": 5,
    "openItems": 6,
    "viewVersions": 7,
    "deleteVersions": 8,
    "cancelCheckout": 9,
    "managePersonalViews": 10,
    "manageLists": 12,
    "viewFormPages": 13,
    "anonymousSearchAccessList": 14,
    "open": 17,
    "viewPages": 18,
    "addAndCustomizePages": 19,
    "applyThemeAndBorder": 20,
    "applyStyleSheets": 21,
    "viewUsageData": 22,
    "createSSCSite": 23,
    "manageSubwebs": 24,
    "createGroups": 25,
    "managePermissions": 26,
    "browseDirectories": 27,
    "browseUserInfo": 28,
    "addDelPrivateWebParts": 29,
    "updatePersonalWebParts": 30,
    "manageWeb": 31,
    "anonymousSearchAccessWebLists": 32,
    "useClientIntegration": 37,
    "useRemoteAPIs": 38,
    "manageAlerts": 39,
    "createAlerts": 40,
    "editMyUserInfo": 41,
    "enumeratePermissions": 63,
    "fullMask": 65
}

function hasPermission(permMask, permLevel){
  if(permLevel === 65){
    return (permMask.High & 32767) === 32767 && permMask.Low === 65535;
  }
  var numericVal = permLevel - 1;
  var indexer = 1;
  if(numericVal > 0 && numericVal < 32){
    indexer = indexer << numericVal;
    return 0 !== (permMask.Low & indexer);
  }
  else if(numericVal >= 32 && numericVal < 64){
    indexer = indexer << numericVal - 32
    return 0 !== (permMask.High & indexer)
  }
  return false;
}

async function getBasePermissions(site, userEmail){
  var url = `${site}/_api/web/getUserEffectivePermissions('${encodeURIComponent("i:0#.f|membership|" + userEmail)}')`
  var effectivePermMaskResp = await get(url);
  var effectivePermMask = await effectivePermMaskResp.json();
  var permissions = [];
  for(var permLevelName in permissionKind){
    var hasPermissionLevel = hasPermission(effectivePermMask, permissionKind[permLevelName]);
    if(hasPermissionLevel){
      permissions.push(permLevelName)
    }
  }
  return permissions;
}
