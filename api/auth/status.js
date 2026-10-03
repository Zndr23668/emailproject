export default function handler(req,res){
  const cookies=Object.fromEntries((req.headers.cookie||"").split(";").filter(Boolean).map(v=>{const i=v.indexOf("=");return [v.slice(0,i).trim(),decodeURIComponent(v.slice(i+1))]}));
  if(!cookies.rpmailclean_session) return res.status(200).json({connected:false});
  try{const s=JSON.parse(Buffer.from(cookies.rpmailclean_session,"base64url").toString());res.status(200).json({connected:true,user:s.user||null,expiresAt:s.expires_at||null})}
  catch{res.status(200).json({connected:false})}
}