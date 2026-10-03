export default async function handler(req,res){
 if(req.method!=="POST")return res.status(405).json({error:"Method not allowed"});
 const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
 if(!cookies.rpmailclean_session)return res.status(401).json({error:"Outlook is not connected."});
 let s;try{s=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString())}catch{return res.status(401).json({error:"Invalid session."})}
 let body;try{body=typeof req.body==="string"?JSON.parse(req.body):req.body}catch{return res.status(400).json({error:"Invalid JSON."})}
 const ids=Array.isArray(body?.messageIds)?body.messageIds.filter(Boolean).slice(0,50):[];const action=body?.action;
 if(!ids.length||!["archive","markRead","markUnread","quarantine","delete","blockSender"].includes(action))return res.status(400).json({error:"Choose messages and a valid action."});
 async function graph(url,opts={}){return fetch(url,{...opts,headers:{Authorization:"Bearer "+s.access_token,"Content-Type":"application/json",...(opts.headers||{})}})}
 async function act(id){
  if(action==="delete")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+id,{method:"DELETE"})).ok;
  if(action==="markRead"||action==="markUnread")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+id,{method:"PATCH",body:JSON.stringify({isRead:action==="markRead"})})).ok;
  if(action==="archive")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+id+"/move",{method:"POST",body:JSON.stringify({destinationId:"archive"})})).ok;
  if(action==="quarantine")return (await graph("https://graph.microsoft.com/v1.0/me/messages/"+id,{method:"PATCH",body:JSON.stringify({isRead:false,categories:["RPMailClean-Quarantine"]})})).ok;
  return false;
 }
 if(action==="blockSender"){
  const senders=Array.isArray(body?.senders)?body.senders.filter(Boolean).slice(0,20):[];
  if(!senders.length)return res.status(400).json({error:"No sender addresses supplied."});
  const current=await graph("https://graph.microsoft.com/v1.0/me/mailboxSettings/junkEmailConfiguration");
  const cfg=await current.json();
  if(!current.ok)return res.status(current.status).json({error:cfg.error?.message||"Could not read Junk Email settings."});
  const blocked=[...new Set([...(cfg.blockedSendersAndDomains||[]),...senders.map(x=>x.toLowerCase())])];
  const updated=await graph("https://graph.microsoft.com/v1.0/me/mailboxSettings/junkEmailConfiguration",{method:"PATCH",body:JSON.stringify({blockedSendersAndDomains:blocked})});
  const out=await updated.json().catch(()=>({}));
  if(!updated.ok)return res.status(updated.status).json({error:out.error?.message||"Could not block sender. Check MailboxSettings.ReadWrite permission."});
  return res.status(200).json({action,blocked:senders});
 }
 const results=await Promise.all(ids.map(async id=>({id,ok:await act(id)})));res.status(200).json({action,results});
}