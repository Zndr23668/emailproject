import { createHmac, randomUUID } from "node:crypto";

export default async function handler(req,res){
  if(req.method!=="GET") return res.status(405).json({error:"Method not allowed"});
  const clientId=process.env.MICROSOFT_CLIENT_ID;
  const clientSecret=process.env.MICROSOFT_CLIENT_SECRET;
  const redirectUri=process.env.MICROSOFT_REDIRECT_URI;
  if(!clientId||!clientSecret||!redirectUri) return res.status(500).json({error:"Microsoft OAuth is not configured."});
  const payload=Buffer.from(JSON.stringify({nonce:randomUUID(),exp:Date.now()+600000})).toString("base64url");
  const sig=createHmac("sha256",clientSecret).update(payload).digest("base64url");
  const state=payload+"."+sig;
  const scope="openid profile email offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite";
  const params=new URLSearchParams({client_id:clientId,response_type:"code",redirect_uri:redirectUri,response_mode:"query",scope,state});
  res.redirect(302,"https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?"+params.toString());
}