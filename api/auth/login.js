export default async function handler(req,res){
  if(req.method!=="GET") return res.status(405).json({error:"Method not allowed"});
  const clientId=process.env.MICROSOFT_CLIENT_ID;
  const redirectUri=process.env.MICROSOFT_REDIRECT_URI;
  if(!clientId||!redirectUri) return res.status(500).json({error:"Microsoft OAuth is not configured. Set MICROSOFT_CLIENT_ID and MICROSOFT_REDIRECT_URI."});
  const state=crypto.randomUUID();
  const scope="openid profile email offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite";
  res.setHeader("Set-Cookie",`rpmailclean_oauth_state=${state}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=600`);
  const params=new URLSearchParams({client_id:clientId,response_type:"code",redirect_uri:redirectUri,response_mode:"query",scope,state});
  res.redirect(302,"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?"+params.toString());
}