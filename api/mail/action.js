export default async function handler(req,res){
 if(req.method!=="POST")return res.status(405).json({error:"Method not allowed"});
 const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
 if(!cookies.rpmailclean_session)return res.status(401).json({error:"Outlook is not connected."});
 let s;try{s=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString())}catch{return res.status(401).json({error:"Invalid session."})}
 let body;try{body=typeof req.body==="string"?JSON.parse(req.body):req.body}catch{return res.status(400).json({error:"Invalid JSON."})}
 const ids=Array.isArray(body?.messageIds)?body.messageIds.filter(Boolean).slice(0,50):[];const action=body?.action;
 if(!ids.length||!["archive","markRead","markUnread","quarantine"].includes(action))return res.status(400).json({error:"Choose messages and a valid action."});
 async function call(id,method,url,payload){const r=await fetch(url,{method,headers:{Authorization:"Bearer "+s.access_token,"Content-Type":"application/json"},body:payload?JSON.stringify(payload):undefined});return r.ok}
 async function act(id){
  if(action==="markRead"||action==="markUnread")return call(id,"PATCH",`https://graph.microsoft.com/v1.0/me/messages/${id}`,{isRead:action==="markRead"});
  if(action==="archive")return call(id,"POST",`https://graph.microsoft.com/v1.0/me/messages/${id}/move`,{destinationId:"archive"});
  return call(id,"PATCH",`https://graph.microsoft.com/v1.0/me/messages/${id}`,{isRead:false,categories:["RPMailClean-Quarantine"]});
 }
 const results=await Promise.all(ids.map(async id=>({id,ok:await act(id)})));res.status(200).json({action,results});
}