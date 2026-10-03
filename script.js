const msgs=[["MW","Marketing Weekly","The 7 things you missed this week","Low"],["MS","Microsoft Security","Unusual sign-in attempt detected","High"],["AP","Adobe Promotions","Your monthly creative roundup","Low"],["VC","Vendor Cloud","Your storage report is ready","Low"]];
const list=document.getElementById("messages");list.innerHTML=msgs.map((m,i)=>`<div class="message"><div class="mailIcon">${m[0]}</div><div><strong>${m[1]}</strong><small>${m[2]}</small></div><span class="risk ${m[3]==="High"?"high":""}">${m[3]}</span></div>`).join("");
const toast=document.getElementById("toast");function notify(t){toast.textContent=t;toast.classList.add("show");setTimeout(()=>toast.classList.remove("show"),2600)}
document.querySelectorAll(".nav[data-view],.textBtn").forEach(b=>b.addEventListener("click",()=>{document.querySelectorAll(".nav").forEach(x=>x.classList.remove("active"));const v=b.dataset.view;if(v){const n=document.querySelector('.nav[data-view="'+v+'"]');if(n)n.classList.add("active");document.getElementById("title").textContent=v==="overview"?"Inbox overview":v.replace(/(^|_)(.)/g,(_,a,c)=>" "+c.toUpperCase()).trim()}notify("View loaded — demo mode")}));
document.getElementById("connectBtn").onclick=()=>notify("Outlook connection ready — Microsoft Graph OAuth will be enabled here.");
document.getElementById("scanBtn").onclick=()=>notify("Inbox scan started — 1,842 messages queued for analysis.");
document.getElementById("cleanupBtn").onclick=()=>notify("Safe cleanup preview created. Nothing was deleted.");
document.getElementById("startBtn").onclick=()=>notify("Cleanup assistant opened.");
document.querySelectorAll("[data-action]").forEach(b=>b.onclick=()=>notify("Review queue opened for "+b.dataset.action+"."));
