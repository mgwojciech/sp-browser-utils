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
    fetch(api,{
        method:"GET",
        headers:{
            accept: "application/json",
            'content-type': "application/json;odata=verbose"
        }
    })
}

function getVerbose(api){
    fetch(api,{
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
