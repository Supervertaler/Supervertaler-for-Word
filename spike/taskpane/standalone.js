// Stand-in for Office.js so the pane can be opened in a plain browser:
//   http://localhost:3000/index.html?standalone
// Just enough of the Word API for what index.html uses. The "document" is
// four segments held in memory; clicking a grid row selects a segment.
(function () {
  const NS = "urn:supervertaler:word:1";
  const segs = [
    ["A mashup is a Web application that combines data from one or more sources.", "Een mashup is een webtoepassing die gegevens uit één of meer bronnen combineert."],
    ["The term implies easy, fast integration.", "De term impliceert een eenvoudige, snelle integratie."],
    ["An example of a mashup is the use of cartographic data from a mapping program.", "Een voorbeeld van een mashup is het gebruik van cartografische gegevens uit een kaarttoepassing."],
    ["This creates a new and distinct Web service.", ""],
  ].map(([s, t], i) => ({ id: "seg" + (i + 1), source: s, target: t, status: "draft", origin: t ? "ai" : "", live: t || s, deleted: !!t }));
  const state = { current: null, selText: "", handlers: [], xml: null };

  const esc = s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const buildXml = () => `<sv:project xmlns:sv="${NS}" source_lang="en" target_lang="nl">` +
    segs.map(s => `<sv:segment id="${s.id}" status="${s.status}" origin="${s.origin}" match="0"><sv:source>${esc(s.source)}</sv:source><sv:target>${esc(s.target)}</sv:target></sv:segment>`).join("") + "</sv:project>";
  state.xml = buildXml();

  const noop = () => {};
  const cc = s => ({
    tag: "sv:seg:" + s.id, title: s.status, text: s.live, load: noop,
    select() { state.current = s.id; setTimeout(fire, 0); },
    insertText(text, where) { if (where === "Replace" || !s.deleted) { s.live = text; s.deleted = true; } else { s.live = text; } },
    getRange() { return { getTrackedChanges: () => ({ items: s.deleted ? [{ type: "Deleted" }, { type: "Added", getRange: () => ({ delete: noop }) }] : [], load: noop }) }; },
    set title(v) { s.status = v; }, get title() { return s.status; },
  });
  const fire = () => state.handlers.forEach(h => h());

  window.Word = {
    ChangeTrackingMode: { trackAll: "TrackAll" },
    run: async fn => {
      const cur = segs.find(x => x.id === state.current);
      const ctx = {
        sync: async () => {},
        document: {
          changeTrackingMode: null,
          customXmlParts: {
            getByNamespace: () => ({ items: [{ getXml: () => ({ value: state.xml }), delete: noop }], load: noop }),
            add: xml => { state.xml = xml; },
          },
          contentControls: {
            items: segs.map(cc), load: noop,
            getByTag: tag => ({ items: segs.filter(s => "sv:seg:" + s.id === tag).map(cc), load: noop }),
          },
          getSelection: () => ({
            text: state.selText, load: noop,
            parentContentControlOrNullObject: cur ? cc(cur) : { isNullObject: true, load: noop, tag: "" },
            contentControls: { items: [], load: noop },
          }),
        },
      };
      if (cur) ctx.document.getSelection().parentContentControlOrNullObject.isNullObject = false;
      return fn(ctx);
    },
  };
  window.Office = {
    HostType: { Word: "Word" },
    EventType: { DocumentSelectionChanged: "sel" },
    onReady: fn => setTimeout(() => fn({ host: "Word" }), 0),
    context: { document: { addHandlerAsync: (_t, h) => state.handlers.push(h) } },
  };
  // dev helpers: window.__word.select("seg2"), window.__word.type("seg2", "new text"), window.__word.selectText("word")
  window.__word = {
    select: id => { state.current = id; fire(); },
    type: (id, text) => { const s = segs.find(x => x.id === id); s.live = text; },
    selectText: t => { state.selText = t; fire(); },
    segs,
  };
})();
