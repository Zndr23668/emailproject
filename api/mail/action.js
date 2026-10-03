export default async function handler(req,res){
 if(req.method!=="POST")return res.status(405).json({error:"Method not allowed"});
 const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
 if(!cookies.rpmailclean_session)return res.status(401).json({error:"Outlook is not connected."});
 let s;try{s=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString())}catch{return res.status(401).json({error:"Invalid session."})}
 let body;try{body=typeof req.body==="string"?JSON.parse(req.body):req.body}catch{return res.status(400).json({error:"Invalid JSON."})}
 const ids=Array.isArray(body?.messageIds)?body.messageIds.filter(Boolean).slice(0,50):[];const action=body?.action;
 if(!ids.length||!["archive","markRead","markUnread","quarantine","delete","blockSender"].includes(action))return res.status(400).json({error:"Choose messages and a valid action."});
 async function graph(url,opts={}){return fetch(url,{...opts,headers:{Authorization:"Bearer "+s.access_token,"Content-Type":"application/json",...(opts.headers||{})}})}
 async function jsonGraph(url,opts={}){const r=await graph(url,opts);const data=await r.json().catch(()=>({}));return {r,data}}
 async function act(id){
  if(action==="delete")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+encodeURIComponent(id),{method:"DELETE"})).ok;
  if(action==="markRead"||action==="markUnread")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+encodeURIComponent(id),{method:"PATCH",body:JSON.stringify({isRead:action==="markRead"})})).ok;
  if(action==="archive")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+encodeURIComponent(id)+"/move",{method:"POST",body:JSON.stringify({destinationId:"archive"})})).ok;
  if(action==="quarantine")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+encodeURIComponent(id),{method:"PATCH",body:JSON.stringify({isRead:false,categories:["RPMailClean-Quarantine"]})})).ok;
  return false;
 }
 if(action==="blockSender"){
  const senders=[...new Set((Array.isArray(body?.senders)?body.senders:[]).filter(Boolean).map(x=>String(x).trim().toLowerCase()))].slice(0,20);
  if(!senders.length)return res.status(400).json({error:"No sender addresses supplied."});
  // Microsoft Graph does not expose /me/mailboxSettings/junkEmailConfiguration.
  // Use the supported Inbox messageRules API to permanently route future mail from
  // each sender to Junk Email, and move the currently selected messages there too.
  const listed=await jsonGraph("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules");
  if(!listed.r.ok)return res.status(listed.r.status).json({error:listed.data?.error?.message||"Could not read Outlook inbox rules. Reconnect Outlook and grant MailboxSettings.ReadWrite."});
  const rules=Array.isArray(listed.data?.value)?listed.data.value:[];
  const created=[];const existing=[];
  for(const sender of senders){
   const key=sender.toLowerCase();
   const found=rules.find(x=>x.isEnabled!==false&&Array.isArray(x.conditions?.senderContains)&&x.conditions.senderContains.some(v=>String(v).toLowerCase()===key));
   if(found){existing.push(sender);continue}
   const maxSeq=rules.reduce((n,x)=>Math.max(n,Number(x.sequence)||0),0);
   const made=await jsonGraph("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",{method:"POST",body:JSON.stringify({
    displayName:"RPMailClean - Block "+sender,
    sequence:maxSeq+created.length+1,
    isEnabled:true,
    conditions:{senderContains:[sender]},
    actions:{moveToFolder:"junkemail",stopProcessingRules:true}
   })});
   if(!made.r.ok)return res.status(made.r.status).json({error:made.data?.error?.message||("Could not create block rule for "+sender+".")});
   created.push(sender);
  }
  const results=await Promise.all(ids.filter(id=>id!=="sender-control").map(async id=>({id,ok:(await graph("https://graph.microsoft.com/v1.0/me/messages/"+encodeURIComponent(id)+"/move",{method:"POST",body:JSON.stringify({destinationId:"junkemail"})})).ok})));
  return res.status(200).json({action,blocked:[...new Set([...created,...existing])],created,existing,results});
 }
 const results=await Promise.all(ids.map(async id=>({id,ok:await act(id)})));res.status(200).json({action,results});
}