export default async function handler(req,res){
  if(req.method!=="GET") return res.status(405).send("Method not allowed");
  const {code,state}=req.query;
  const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
  if(!code||!state||state!==cookies.rpmailclean_oauth_state) return res.status(400).send("Invalid OAuth state.");
  const clientId=process.env.MICROSOFT_CLIENT_ID, clientSecret=process.env.MICROSOFT_CLIENT_SECRET, redirectUri=process.env.MICROSOFT_REDIRECT_URI;
  if(!clientId||!clientSecret||!redirectUri) return res.status(500).send("Microsoft OAuth is not configured.");
  const token=await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token",{method:"POST",headers:{"content-type":"application/x-www-form-urlencoded"},body:new URLSearchParams({client_id:clientId,client_secret:clientSecret,grant_type:"authorization_code",code,redirect_uri:redirectUri,scope:"openid profile email offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite"})});
  const data=await token.json();
  if(!token.ok||!data.access_token) return res.status(502).send("Microsoft authorization failed.");
  const me=await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName",{headers:{Authorization:"Bearer "+data.access_token}});
  const user=await me.json();
  if(!me.ok) return res.status(502).send("Could not read the Microsoft account.");
  const session=Buffer.from(JSON.stringify({access_token:data.access_token,refresh_token:data.refresh_token||null,expires_at:Date.now()+((data.expires_in||3600)*1000),user})).toString("base64url");
  res.setHeader("Set-Cookie",[
    `rpmailclean_session=${session}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=3600`,
    "rpmailclean_oauth_state=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0"
  ]);
  res.redirect(302,"/?connected=1");
}