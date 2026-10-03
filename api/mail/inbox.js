export default async function handler(req,res){
 if(req.method!=="GET")return res.status(405).json({error:"Method not allowed"});
 const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
 if(!cookies.rpmailclean_session)return res.status(401).json({error:"Outlook is not connected."});
 let s;try{s=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString())}catch{return res.status(401).json({error:"Invalid session."})}
 const top=Math.min(Math.max(Number(req.query.top)||50,1),100);
 const url="https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top="+top+"&$orderby=receivedDateTime%20desc&$select=id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview,categories";
 const r=await fetch(url,{headers:{Authorization:"Bearer "+s.access_token}});
 const d=await r.json();if(!r.ok)return res.status(r.status).json({error:d.error?.message||"Microsoft Graph request failed."});
 const messages=(d.value||[]).map(m=>{const text=((m.subject||"")+" "+(m.bodyPreview||"")).toLowerCase(),sender=(m.from?.emailAddress?.address||"").toLowerCase();let risk="Low";let reason="";
 if(/password|verify your account|urgent|wire transfer|gift card|crypto|invoice attached/.test(text)||/noreply@.*\.ru$|@.*\.xyz$/.test(sender)){risk="High";reason="Contains common phishing or suspicious-mail signals."}
 else if(/unsubscribe|newsletter|promotion|offer|sale|weekly roundup/.test(text)){risk="Review";reason="Looks like bulk or promotional mail."}
 return {...m,risk,reason}});
 res.status(200).json({connected:true,user:s.user,messages});
}