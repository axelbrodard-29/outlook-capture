/* Capture d'e-mail – complément Outlook
 * 1. Lit l'e-mail ouvert (en-tête + corps HTML + images intégrées)
 * 2. L'affiche dans une zone invisible à largeur fixe et le convertit en image (html2canvas)
 * 3. Demande le format : PNG, JPEG, PDF ou presse-papiers
 */
(function () {
  "use strict";

  var RENDER_WIDTH = 820;          // largeur de la capture en pixels CSS
  var MAX_CANVAS_HEIGHT = 30000;   // limite de sécurité du navigateur
  var current = null;              // { canvas, fileBase, previewUrl }

  var $ = function (id) { return document.getElementById(id); };

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      showError("Ce complément fonctionne uniquement dans Outlook.");
      return;
    }
    document.querySelectorAll(".fmt").forEach(function (btn) {
      btn.addEventListener("click", function () { exportAs(btn.dataset.format); });
    });
    $("recapture").addEventListener("click", capture);

    // Si le volet reste ouvert et que l'on change d'e-mail, on recapture.
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, capture);
    } catch (e) { /* non supporté : sans importance */ }

    capture();
  });

  /* ---------- Étape 1 : lecture de l'e-mail ---------- */

  function getBodyHtml(item) {
    return new Promise(function (resolve, reject) {
      item.body.getAsync(Office.CoercionType.Html, function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value || "");
        else reject(new Error("Impossible de lire le contenu de l'e-mail : " + (res.error && res.error.message)));
      });
    });
  }

  function getAttachmentBase64(item, id) {
    return new Promise(function (resolve) {
      item.getAttachmentContentAsync(id, function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded &&
            res.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          resolve(res.value.content);
        } else {
          resolve(null);
        }
      });
    });
  }

  function guessMime(att) {
    if (att.contentType && /^image\//i.test(att.contentType)) return att.contentType;
    var ext = (att.name || "").split(".").pop().toLowerCase();
    return { jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp" }[ext] || "image/png";
  }

  // Remplace les images intégrées (src="cid:...") par leur contenu réel.
  async function inlineCidImages(item, html) {
    var atts = item.attachments || [];
    if (typeof item.getAttachmentContentAsync !== "function") return html;
    for (var i = 0; i < atts.length; i++) {
      var att = atts[i];
      if (!att.isInline || !att.contentId) continue;
      var b64 = await getAttachmentBase64(item, att.id);
      if (!b64) continue;
      var cid = att.contentId.replace(/^<|>$/g, "");
      var escaped = cid.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      html = html.replace(new RegExp("cid:" + escaped, "gi"), "data:" + guessMime(att) + ";base64," + b64);
    }
    return html;
  }

  /* ---------- Étape 2 : mise en page et rendu ---------- */

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function fmtPeople(list) {
    if (!list) return "";
    if (!Array.isArray(list)) list = [list];
    return list.map(function (p) {
      if (!p) return "";
      var name = p.displayName || "";
      var mail = p.emailAddress || "";
      return name && mail && name !== mail ? name + " <" + mail + ">" : (name || mail);
    }).filter(Boolean).join("; ");
  }

  function fmtDate(d) {
    if (!d) return "";
    d = new Date(d);
    return d.toLocaleString("fr-CH", {
      weekday: "long", day: "numeric", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit"
    });
  }

  function buildHeader(item) {
    var rows = [
      ["De", fmtPeople(item.from)],
      ["Envoyé", fmtDate(item.dateTimeCreated)],
      ["À", fmtPeople(item.to)],
      ["Cc", fmtPeople(item.cc)]
    ];
    var files = (item.attachments || []).filter(function (a) { return !a.isInline; })
      .map(function (a) { return a.name; });
    if (files.length) rows.push(["Pièces jointes", files.join(", ")]);

    var html = '<div id="__cap-header" style="font-family:\'Segoe UI\',Arial,sans-serif;color:#242424;' +
      'padding:20px 24px 14px;border-bottom:1px solid #d1d1d1;margin-bottom:16px;background:#fff">' +
      '<div style="font-size:20px;font-weight:600;margin-bottom:10px;line-height:1.3">' +
      esc(item.subject || "(Sans objet)") + "</div>" +
      '<table style="border-collapse:collapse;font-size:13px;line-height:1.5">';
    rows.forEach(function (r) {
      if (!r[1]) return;
      html += '<tr><td style="color:#616161;padding:1px 14px 1px 0;vertical-align:top;white-space:nowrap">' +
        esc(r[0]) + '</td><td style="padding:1px 0">' + esc(r[1]) + "</td></tr>";
    });
    return html + "</table></div>";
  }

  function buildDocument(item, bodyHtml) {
    var parsed = new DOMParser().parseFromString(bodyHtml, "text/html");
    var styles = Array.prototype.map.call(parsed.querySelectorAll("style"), function (s) {
      return s.textContent;
    }).join("\n");
    var clean = DOMPurify.sanitize(parsed.body ? parsed.body.innerHTML : bodyHtml, {
      FORBID_TAGS: ["script", "iframe", "object", "embed", "form", "input", "button", "textarea", "select", "link", "meta", "base"],
      FORBID_ATTR: ["srcset"]
    });
    return "<!DOCTYPE html><html><head><meta charset='utf-8'><style>" +
      "html,body{margin:0;padding:0;background:#fff;width:" + RENDER_WIDTH + "px}" +
      "#__cap-body{padding:0 24px 24px;font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:14.5px;" +
      "color:#000;overflow-wrap:anywhere}" +
      "#__cap-body img{max-width:100%;height:auto}" +
      "</style><style>" + styles.replace(/<\/style/gi, "") + "</style></head><body>" +
      buildHeader(item) + '<div id="__cap-body">' + clean + "</div></body></html>";
  }

  function loadFrame(html) {
    var frame = $("render-frame");
    return new Promise(function (resolve) {
      frame.onload = function () { resolve(frame); };
      frame.srcdoc = html;
    });
  }

  function waitForImages(doc, timeoutMs) {
    var imgs = Array.prototype.slice.call(doc.images);
    var pending = imgs.filter(function (img) { return !img.complete; }).map(function (img) {
      return new Promise(function (r) { img.addEventListener("load", r); img.addEventListener("error", r); });
    });
    return Promise.race([
      Promise.all(pending),
      new Promise(function (r) { setTimeout(r, timeoutMs); })
    ]);
  }

  async function render(html) {
    var frame = await loadFrame(html);
    var doc = frame.contentDocument;
    await waitForImages(doc, 8000);
    if (doc.fonts && doc.fonts.ready) { try { await doc.fonts.ready; } catch (e) {} }

    var height = Math.max(doc.body.scrollHeight, doc.documentElement.scrollHeight);
    frame.style.height = height + "px";

    var scale = Math.min(2, MAX_CANVAS_HEIGHT / height);
    return html2canvas(doc.body, {
      backgroundColor: "#ffffff",
      scale: Math.max(scale, 0.5),
      useCORS: true,
      allowTaint: false,
      imageTimeout: 8000,
      logging: false,
      width: RENDER_WIDTH,
      height: height,
      windowWidth: RENDER_WIDTH,
      windowHeight: height
    });
  }

  function fileBaseName(item) {
    // Nom de fichier en ASCII simple (accents retirés) pour éviter les refus de téléchargement.
    var subject = (item.subject || "e-mail").normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[–—]/g, "-").replace(/[^\x20-\x7e]+/g, "")
      .replace(/[\\/:*?"<>|]+/g, " ")
      .replace(/\s+/g, " ").trim().slice(0, 80) || "e-mail";
    var d = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
    var p = function (n) { return String(n).padStart(2, "0"); };
    return subject + "_" + d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
  }

  async function capture() {
    var item = Office.context.mailbox.item;
    showStep("loading");
    hideStatus();
    if (!item) { showError("Ouvrez ou sélectionnez un e-mail, puis réessayez."); return; }
    try {
      setLoading("Lecture de l'e-mail…");
      var body = await getBodyHtml(item);
      setLoading("Récupération des images…");
      body = await inlineCidImages(item, body);
      setLoading("Création de la capture…");
      var canvas = await render(buildDocument(item, body));
      var blob = await toBlob(canvas, "image/png");

      if (current && current.previewUrl) URL.revokeObjectURL(current.previewUrl);
      current = { canvas: canvas, fileBase: fileBaseName(item), previewUrl: URL.createObjectURL(blob) };
      $("preview").src = current.previewUrl;
      showStep("choose");
    } catch (err) {
      console.error(err);
      showError("La capture a échoué : " + (err && err.message ? err.message : err));
    }
  }

  /* ---------- Étape 3 : export ---------- */

  function toBlob(canvas, type, quality) {
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (b) { b ? resolve(b) : reject(new Error("conversion impossible")); }, type, quality);
    });
  }

  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
  }

  function buildPdf(canvas) {
    var jsPDF = window.jspdf.jsPDF;
    var pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait", compress: true });
    var margin = 10;
    var pageW = 210 - margin * 2;
    var pageH = 297 - margin * 2;
    var pxPerMm = canvas.width / pageW;
    var slicePx = Math.floor(pageH * pxPerMm);

    for (var y = 0, page = 0; y < canvas.height; y += slicePx, page++) {
      var h = Math.min(slicePx, canvas.height - y);
      var slice = document.createElement("canvas");
      slice.width = canvas.width;
      slice.height = h;
      var ctx = slice.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, slice.width, h);
      ctx.drawImage(canvas, 0, y, canvas.width, h, 0, 0, canvas.width, h);
      if (page > 0) pdf.addPage();
      pdf.addImage(slice.toDataURL("image/jpeg", 0.95), "JPEG", margin, margin, pageW, h / pxPerMm);
    }
    return pdf.output("blob");
  }

  async function exportAs(format) {
    if (!current) return;
    setButtonsDisabled(true);
    hideStatus();
    try {
      var base = current.fileBase;
      if (format === "png") {
        download(await toBlob(current.canvas, "image/png"), base + ".png");
        showOk("PNG enregistré : " + base + ".png");
      } else if (format === "jpeg") {
        download(await toBlob(current.canvas, "image/jpeg", 0.92), base + ".jpg");
        showOk("JPEG enregistré : " + base + ".jpg");
      } else if (format === "pdf") {
        download(buildPdf(current.canvas), base + ".pdf");
        showOk("PDF enregistré : " + base + ".pdf");
      } else if (format === "clipboard") {
        if (!navigator.clipboard || typeof ClipboardItem === "undefined") {
          throw new Error("le presse-papiers n'est pas accessible dans cette version d'Outlook.");
        }
        var blob = await toBlob(current.canvas, "image/png");
        await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        showOk("Capture copiée. Collez-la avec Ctrl+V.");
        return;
      }
      $("fallback").hidden = false;
    } catch (err) {
      console.error(err);
      showError("Échec : " + (err && err.message ? err.message : err));
    } finally {
      setButtonsDisabled(false);
    }
  }

  /* ---------- Interface ---------- */

  function showStep(name) {
    $("step-loading").hidden = name !== "loading";
    $("step-choose").hidden = name !== "choose";
    $("fallback").hidden = true;
  }
  function setLoading(t) { $("loading-text").textContent = t; }
  function setButtonsDisabled(v) {
    document.querySelectorAll(".fmt").forEach(function (b) { b.disabled = v; });
  }
  function hideStatus() { $("status").hidden = true; }
  function showOk(t) { var s = $("status"); s.textContent = t; s.className = "ok"; s.hidden = false; }
  function showError(t) {
    $("step-loading").hidden = true;
    var s = $("status"); s.textContent = t; s.className = "err"; s.hidden = false;
  }
})();
