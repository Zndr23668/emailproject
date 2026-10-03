export default async function handler(req,res){
  if(req.method!=="GET") return res.status(405).json({error:"Method not allowed"});
  const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
  if(!cookies.rpmailclean_session) return res.status(401).json({error:"Outlook is not connected."});
  let session; try{session=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString())}catch{return res.status(401).json({error:"Invalid session."})}
  if(!session.access_token) return res.status(401).json({error:"Missing access token."});
  const top=Math.min(Math.max(Number(req.query.top)||25,1),100);
  const url="https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top="+top+"&$select=id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview";
  const r=await fetch(url,{headers:{Authorization:"Bearer "+session.access_token}});
  const data=await r.json();
  if(!r.ok) return res.status(r.status).json({error:data.error?.message||"Microsoft Graph request failed."});
  res.status(200).json({connected:true,user:session.user,messages:data.value||[]});
}