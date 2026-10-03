export default async function handler(req,res){
 if(req.method!=="POST") return res.status(405).json({error:"Method not allowed"});
 const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
 if(!cookies.rpmailclean_session)return res.status(401).json({error:"Outlook is not connected."});
 let s;try{s=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString())}catch{return res.status(401).json({error:"Invalid session."})}
 let body;try{body=typeof req.body==="string"?JSON.parse(req.body):req.body}catch{return res.status(400).json({error:"Invalid JSON."})}
 const ids=Array.isArray(body?.messageIds)?body.messageIds.filter(Boolean).slice(0,50):[];
 const action=body?.action;
 const allowed=["archive","markRead","markUnread","deleteToRecovery"];
 if(!ids.length||!allowed.includes(action))return res.status(400).json({error:"Choose messages and a valid reversible action."});
 const graphAction=action==="archive"?async id=>{const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`,{method:"PATCH",headers:{Authorization:"Bearer "+s.access_token,"Content-Type":"application/json"},body:JSON.stringify({parentFolderId:"archive"})});return r.ok}:action==="markRead"||action==="markUnread"?async id=>{const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`,{method:"PATCH",headers:{Authorization:"Bearer "+s.access_token,"Content-Type":"application/json"},body:JSON.stringify({isRead:action==="markRead"})});return r.ok}:async id=>{const r=await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`,{method:"PATCH",headers:{Authorization:"Bearer "+s.access_token,"Content-Type":"application/json"},body:JSON.stringify({isRead:false,categories:["RPMailClean-Recovery"]})});return r.ok};
 const results=await Promise.all(ids.map(async id=>({id,ok:await graphAction(id)})));
 res.status(200).json({action,results});
}