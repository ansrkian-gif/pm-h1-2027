(function () {
  var pin = window.DASHBOARD_PIN || "";
  if (!pin) return;
  if (sessionStorage.getItem("pm-pin") === pin) return;
  var wrap = document.createElement("div");
  wrap.setAttribute("id", "pinGate");
  wrap.style.cssText =
    "position:fixed;inset:0;z-index:99999;background:#0f3d3e;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;color:#fff;font-family:system-ui,sans-serif;padding:max(1rem,env(safe-area-inset-top)) max(1rem,env(safe-area-inset-right)) max(1.4rem,env(safe-area-inset-bottom)) max(1rem,env(safe-area-inset-left));box-sizing:border-box";
  wrap.innerHTML =
    '<div style="font-weight:800;font-size:clamp(1.2rem,5vw,1.5rem);text-align:center;letter-spacing:-0.02em">PM Nov 2027</div>' +
    '<div style="opacity:.8;font-size:.95rem;text-align:center">Same PIN as PM H1 2027</div>' +
    '<input id="pinIn" type="password" inputmode="numeric" maxlength="8" placeholder="PIN" autocomplete="off" style="font-size:1.25rem;padding:14px 16px;border-radius:12px;border:0;width:min(220px,80vw);text-align:center">' +
    '<button id="pinGo" type="button" style="padding:14px 32px;border:0;border-radius:999px;font-weight:800;font-size:1rem;cursor:pointer">Open</button>';
  document.documentElement.appendChild(wrap);
  function go() {
    var v = document.getElementById("pinIn").value;
    if (v === pin) {
      sessionStorage.setItem("pm-pin", v);
      wrap.remove();
    } else {
      document.getElementById("pinIn").style.outline = "2px solid #fca5a5";
    }
  }
  document.getElementById("pinGo").onclick = go;
  document.getElementById("pinIn").addEventListener("keydown", function (e) {
    if (e.key === "Enter") go();
  });
  setTimeout(function () {
    var el = document.getElementById("pinIn");
    if (el) el.focus();
  }, 50);
})();
