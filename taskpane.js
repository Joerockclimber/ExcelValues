
Office.onReady(() => {
  const out = document.getElementById("out");
  const status = document.getElementById("status");

  let lastSig = "";
  let busy = false;
  const intervalMs = 250;

  async function tick() {
    if (busy) return;
    busy = true;
    try {
      await Excel.run(async (ctx) => {
        const rng = ctx.workbook.getSelectedRange();
        rng.load(["address", "text"]);
        await ctx.sync();

        const shown = toTSV(rng.text);
        const sig = rng.address + "::" + shown;
        if (sig !== lastSig) {
          lastSig = sig;
          out.textContent = `${shown || "(blank)"}`;
          status.textContent = "Selection Viewer";
        }
      });
    } catch (e) {
      status.textContent = "Viewer running…";
    } finally {
      busy = false;
    }
  }

  function toTSV(arr2d) {
    try {
      return arr2d.map(r => r.map(c => (c == null ? "" : String(c))).join("\t")).join("\n");
    } catch { return ""; }
  }

  tick();
  setInterval(tick, intervalMs);
});
