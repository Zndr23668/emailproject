const sampleMsgs=[["MW","Marketing Weekly","The 7 things you missed this week","Low"],["MS","Microsoft Security","Unusual sign-in attempt detected","High"],["AP","Adobe Promotions","Your monthly creative roundup","Low"],["VC","Vendor Cloud","Your storage report is ready","Low"]];
const list=document.getElementById("messages");const toast=document.getElementById("toast");
function notify(t){toast.textContent=t;toast.classList.add("show");setTimeout(()=>toast.classList.remove("show"),2600)}
function render(messages){list.innerHTML=messages.map(m=>{const sender=m.from?.emailAddress?.name||m[1],subject=m.subject||m[2]||"",preview=m.bodyPreview||"",risk=m.risk||(m[3]||"Low");return `<div class="message"><div class="mailIcon">${(sender||"??").slice(0,2).toUpperCase()}</div><div><strong>${sender}</strong><small>${subject||preview}</small></div><span class="risk ${risk==="High"?"high":""}">${risk}</span></div>`}).join("")}
render(sampleMsgs);
async function checkStatus(){try{const r=await fetch("/api/auth/status");const d=await r.json();if(d.connected){document.getElementById("connectionState").textContent=d.user?.mail||d.user?.userPrincipalName||"Connected";document.getElementById("connectBtn").textContent="Outlook connected";await loadInbox(false)}}catch{}}
async function loadInbox(show=true){try{const r=await fetch("/api/mail/inbox?top=25");const d=await r.json();if(!r.ok)throw new Error(d.error);render(d.messages||[]);if(show)notify(`${d.messages?.length||0} Outlook messages loaded.`)}catch(e){notify(e.message||"Could not load Outlook mail.")}}
document.getElementById("connectBtn").onclick=()=>{window.location.href="/api/auth/login"};
document.getElementById("scanBtn").onclick=()=>loadInbox(true);
document.getElementById("cleanupBtn").onclick=()=>notify("Cleanup review created. Nothing has been changed.");
document.getElementById("startBtn").onclick=()=>notify("Cleanup assistant opened.");
document.querySelectorAll(".nav[data-view],.textBtn").forEach(b=>b.addEventListener("click",()=>{document.querySelectorAll(".nav").forEach(x=>x.classList.remove("active"));const v=b.dataset.view;if(v){const n=document.querySelector('.nav[data-view="'+v+'"]');if(n)n.classList.add("active");document.getElementById("title").textContent=v==="overview"?"Inbox overview":v.replace(/(^|_)(.)/g,(_,a,c)=>" "+c.toUpperCase()).trim()}notify("View loaded")}));
document.querySelectorAll("[data-action]").forEach(b=>b.onclick=()=>notify("Review queue opened for "+b.dataset.action+"."));
checkStatus();
