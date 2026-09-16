"""The local engine, first cut: serves the task pane and answers its questions
from a Supervertaler database, read-only.

    python engine/engine.py --db "D:\\path\\supervertaler.db" --termbases 4 [--port 3000]

Endpoints (all JSON):
    GET  /api/status                      what is loaded
    GET  /api/terms?text=&src=nl&tgt=en   termbase hits for a sentence
    GET  /api/tm?text=&src=nl&tgt=en      TM candidates for a sentence (the pane scores them)
    POST /api/terms  {src, tgt}           add a term: kept in a local pending file, not the database

The database is opened read-only. Nothing here writes to it. Terms added
from the pane go to engine/pending_terms.json (ignored by git) and are served
alongside the database terms until they are imported properly.

Scale: terms for the configured termbases are held in memory (51,000 for
BEIJER, trivial). TM candidates come from the database's own FTS5 index on
translation units, 30 ms on 877,000 units.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(HERE, "..", "spike", "taskpane")
PENDING = os.path.join(HERE, "pending_terms.json")

_word = re.compile(r"[\w'’-]+", re.UNICODE)

# Function words that some termbases carry as entries ("voor -> for"). A term
# made only of these is noise in a lens, whatever the termbase says.
STOP = set("""de het een en van voor dit dat op in te met is zijn aan bij of om als
door uit naar over ook nog dan wordt worden werd hij zij ze het er
the a an and of for this that with to at by on in is are be as or from into""".split())
def trivial(key: str) -> bool:
    ws = key.split()
    return all(w in STOP for w in ws) or (len(ws) == 1 and len(ws[0]) < 3)


def norm(s: str) -> str:
    return " ".join(_word.findall(s.lower()))


class Engine:
    def __init__(self, db_path: str, termbase_ids: list[str]):
        self.db_path = db_path
        self.con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True, check_same_thread=False)
        self.lock = threading.Lock()
        self.termbases = self._load_termbases(termbase_ids)
        self.terms: dict[str, dict] = {}          # norm(source) -> {src, tgts, kind, pair}
        self.max_words = 1
        for tb in self.termbases:
            self._load_terms(tb)
        self._load_pending()
        self.tm_units = self.con.execute("select count(*) from translation_units").fetchone()[0]

    # ---- loading
    def _load_termbases(self, ids):
        rows = self.con.execute(
            "select id, name, source_lang, target_lang from termbases where id in (%s)"
            % ",".join("?" * len(ids)), ids).fetchall()
        return [{"id": r[0], "name": r[1], "src": (r[2] or "").lower(), "tgt": (r[3] or "").lower(), "count": 0} for r in rows]

    def _add_term(self, src, tgt, kind, pair):
        key = norm(src)
        if not key or trivial(key):
            return
        e = self.terms.setdefault(key, {"src": src, "tgts": [], "kind": kind, "pair": pair})
        if tgt and tgt not in e["tgts"]:
            e["tgts"].append(tgt)
        if kind == "nt":
            e["kind"] = "nt"
        self.max_words = max(self.max_words, key.count(" ") + 1)

    def _load_terms(self, tb):
        rows = self.con.execute(
            "select source_term, target_term, is_nontranslatable, synonyms from termbase_terms where termbase_id = ?",
            (str(tb["id"]),)).fetchall()
        pair = (tb["src"], tb["tgt"])
        for src, tgt, nt, syn in rows:
            if not src:
                continue
            self._add_term(src, tgt or "", "nt" if nt else "", pair)
            for s in (syn or "").split(";"):
                if s.strip():
                    self._add_term(src, s.strip(), "", pair)
        tb["count"] = len(rows)

    def _load_pending(self):
        if os.path.exists(PENDING):
            for t in json.load(open(PENDING, encoding="utf-8")):
                self._add_term(t["src"], t["tgt"], "", (t.get("src_lang", ""), t.get("tgt_lang", "")))

    # ---- queries
    def status(self):
        return {"ok": True, "db": os.path.basename(self.db_path), "termbases": self.termbases,
                "terms": len(self.terms), "tm_units": self.tm_units}

    def term_hits(self, text: str, src: str, tgt: str):
        words = _word.findall(text.lower())
        seen, out = set(), []
        for n in range(min(self.max_words, len(words)), 0, -1):
            for i in range(len(words) - n + 1):
                key = " ".join(words[i:i + n])
                e = self.terms.get(key)
                if e and key not in seen and (not e["pair"][0] or e["pair"] == (src, tgt)):
                    seen.add(key)
                    out.append({"src": e["src"], "tgts": e["tgts"], "kind": e["kind"]})
        return out

    def tm_candidates(self, text: str, src: str, tgt: str, limit: int = 40):
        words = sorted({w for w in _word.findall(text.lower()) if len(w) >= 4}, key=len, reverse=True)[:8]
        if not words:
            return []
        q = " OR ".join('"' + w.replace('"', "") + '"' for w in words)
        with self.lock:
            rows = self.con.execute(
                "select t.source_text, t.target_text, t.tm_id from translation_units_fts f "
                "join translation_units t on t.id = f.rowid "
                "where translation_units_fts match ? and lower(substr(t.source_lang,1,2)) = ? "
                "and lower(substr(t.target_lang,1,2)) = ? limit ?", (q, src, tgt, limit)).fetchall()
        return [{"source": s, "target": t, "name": name} for s, t, name in rows]

    def add_pending(self, src: str, tgt: str, src_lang: str, tgt_lang: str):
        items = json.load(open(PENDING, encoding="utf-8")) if os.path.exists(PENDING) else []
        items.append({"src": src, "tgt": tgt, "src_lang": src_lang, "tgt_lang": tgt_lang})
        json.dump(items, open(PENDING, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        self._add_term(src, tgt, "", (src_lang, tgt_lang))


class Handler(SimpleHTTPRequestHandler):
    engine: Engine = None

    def log_message(self, fmt, *args):            # quiet, except errors
        if args and str(args[1]).startswith(("4", "5")):
            super().log_message(fmt, *args)

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        if not u.path.startswith("/api/"):
            return super().do_GET()
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        e = self.engine
        try:
            if u.path == "/api/status":
                return self._json(e.status())
            if u.path == "/api/terms":
                return self._json(e.term_hits(q.get("text", ""), q.get("src", "").lower(), q.get("tgt", "").lower()))
            if u.path == "/api/tm":
                return self._json(e.tm_candidates(q.get("text", ""), q.get("src", "").lower(), q.get("tgt", "").lower()))
        except Exception as ex:                       # noqa: BLE001
            return self._json({"error": str(ex)}, 500)
        self._json({"error": "unknown endpoint"}, 404)

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        data = json.loads(self.rfile.read(n) or b"{}")
        if u.path == "/api/terms":
            src, tgt = (data.get("src") or "").strip(), (data.get("tgt") or "").strip()
            if not src or not tgt:
                return self._json({"error": "src and tgt required"}, 400)
            self.engine.add_pending(src, tgt, (data.get("src_lang") or "").lower(), (data.get("tgt_lang") or "").lower())
            return self._json({"ok": True, "pending": True})
        self._json({"error": "unknown endpoint"}, 404)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", required=True)
    ap.add_argument("--termbases", default="4", help="comma-separated termbase ids")
    ap.add_argument("--port", type=int, default=3000)
    a = ap.parse_args()
    Handler.engine = Engine(a.db, [x.strip() for x in a.termbases.split(",") if x.strip()])
    st = Handler.engine.status()
    print("engine: %s | termbases: %s | %d terms | %d TM units"
          % (st["db"], ", ".join(f'{t["name"]} ({t["count"]})' for t in st["termbases"]), st["terms"], st["tm_units"]))
    srv = ThreadingHTTPServer(("127.0.0.1", a.port), partial(Handler, directory=STATIC))
    print("serving the pane on http://localhost:%d/index.html" % a.port)
    srv.serve_forever()


if __name__ == "__main__":
    main()
